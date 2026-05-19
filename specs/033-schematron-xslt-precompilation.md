# Spec: Schematron-to-XSLT Precompilation for Semantic Validation

## Status

Proposed (May 19, 2026). Implements ADR-004
(`docs/adr/README.md` — "Schematron-to-XSLT Precompilation for
Semantic Validation"). Replaces the per-rule Python evaluator with
a single C-speed XSLT pass for the locally-expressible majority of
the SDK semantic-rule corpus. Keeps imperative Python for rules
that require cross-part `document()` traversal across OOXML parts.

## Problem

`data/openxml/schematrons.json` carries ~948 SDK semantic rules as
`{Context, Test, App}` triples. `codegen/schematron_loader.py`
parses each into one of 14 typed `SchematronType` categories;
`codegen/schematron_bridge.py` lifts those into typed
`SemanticConstraint` instances; the runtime walks every rule
against every matching context node in a Python loop.

The README benchmark records ~2.2× warm latency versus the .NET
SDK on a 798K DOCX. The 1.2× cold gap closes on warmth, which
locates the bottleneck above libxml2 parsing — the per-rule Python
loop is the dominant cost on warm validations. Microsoft's SDK
runs the same rule corpus we ship as JSON, but inside a compiled
.NET pipeline rather than an interpreted Python one.

The rule corpus is Schematron-shaped already (value ranges,
uniqueness, cross-references, co-occurrence, conditional values),
and ISO Schematron has had a standard XSLT-skeleton compiler since
the early 2000s. `lxml.isoschematron` ships the skeleton already.
Compilation is "many minutes for mid-sized rule sets" per the
Schematron.com docs, which is fine — it is a build step, not a
runtime step.

## Why This Matters

- **Warm-path speedup.** A single XSLT pass replaces ~948 Python
  iterations per document. Expected to close most of the 2.2× warm
  gap to .NET; possibly past it on large documents where libxml2's
  pattern matcher amortizes across rules.
- **Compiled XSLT is a portable artifact.** Any libxml2+libxslt
  environment can load it — WebAssembly, Google Apps Script (via a
  WASM stack), browser-side validators — without re-porting the
  rule corpus. The JSON stays the single source of truth; the
  compiled XSLT is the shared execution substrate downstream.
- **Sharper cross-part boundary.** Today the imperative evaluator
  handles every rule type. After this spec it handles only the
  rules that genuinely need `document()` traversal across OOXML
  parts, and can be optimized for those patterns specifically.
- **Error-model parity stays mechanical.** The rule corpus is
  unchanged — only the engine is different. SVRL output adapts
  back to the existing `ValidationError` shape, so downstream
  evidence reporting and SDK-parity comparisons are unaffected.

## Normative References

- `docs/adr/README.md` — ADR-004 Schematron-to-XSLT
  Precompilation. This spec is its execution plan.
- `specs/001-exhaustive-schematron-validation.md` — the SDK
  semantic-rule coverage push that produced the JSON corpus this
  spec compiles.
- `specs/009-bridge-whole-dataset-invariants.md` — invariant
  guards on bridge fidelity. The XSLT pipeline must respect them
  (no silent rule loss between JSON and compiled artifact).
- `specs/026-self-parity-sovereign-gate-roadmap.md` — sovereign
  gates pivot. Faster validation strengthens the self-parity gate
  as a CI signal.
- ISO Schematron skeleton: `iso_dsdl_include.xsl` →
  `iso_abstract_expand.xsl` → `iso_svrl_for_xslt1.xsl` (bundled
  in `lxml.isoschematron`).

## Approach

Three phases. Each phase ships in its own release and leaves the
existing per-rule Python evaluator untouched as the fallback;
cutover to XSLT-first happens at the end of Phase 3.

### Phase 1 — end-to-end spike (smallest slice)

Goal: prove the JSON → `.sch` → XSLT → validation path works on
real fixtures, and measure the speedup on a single rule category
before committing to the full corpus.

1. **`codegen/schematron_compiler.py`**: new module that emits a
   single `.sch` containing every rule of one
   `SchematronType` — pick `ATTRIBUTE_VALUE_RANGE` (the simplest
   typed category, no `document()` calls). Group rules by
   `Context` to keep the schema readable; one
   `<sch:rule context="…">` per distinct context.
2. **Compile via `lxml.isoschematron.Schematron(file=…)`** at
   build time; serialize the compiled XSLT to
   `data/openxml/compiled/range.xslt`. Compilation is invoked by
   a new `scripts/compile_schematron.py` entry point.
3. **Runtime path**: an opt-in `--use-compiled-schematron` CLI
   flag loads the XSLT artifact via `lxml.etree.XSLT` and runs it
   per document. SVRL output is parsed but **not yet** adapted to
   `ValidationError`; instead, the spike emits a side-by-side
   report comparing XSLT findings against the Python evaluator's
   findings for the same rule type.
4. **Benchmark**: extend the README benchmark harness to record
   warm/cold latency for the spike rule type, Python-loop vs
   compiled-XSLT, on the existing 798K DOCX fixture.
5. **Tests** (`tests/test_schematron_compiler.py`): emitted
   `.sch` is well-formed, compiles without error, and produces
   SVRL with at least one assertion firing on a fixture known to
   violate the rule.

### Phase 2 — full locally-expressible coverage

Goal: extend the codegen to every `SchematronType` whose XPath
test does not require `document()` traversal across OOXML parts.

The 14 typed categories split as follows after re-inspecting
`schematron_loader.py`:

| Category | XSLT-compilable | Reason |
|---|---|---|
| `ATTRIBUTE_VALUE_RANGE`     | yes | numeric XPath only |
| `ATTRIBUTE_VALUE_LENGTH`    | yes | `string-length()` only |
| `ATTRIBUTE_VALUE_PATTERN`   | yes | `matches()` only |
| `UNIQUE_ATTRIBUTE`          | yes | `distinct-values()` over local nodeset |
| `ATTRIBUTE_NOT_EQUAL`       | yes | scalar comparison |
| `ATTRIBUTE_EQUALS`          | yes | scalar comparison |
| `ATTRIBUTE_COMPARISON`      | yes | scalar comparison |
| `ATTRIBUTES_PRESENT`        | yes | attribute existence |
| `CONDITIONAL_VALUE`         | yes | attribute existence + equality |
| `OR_CONDITION`              | yes | XPath `or` |
| `AND_CONDITION`             | yes | XPath `and` |
| `RELATIONSHIP_TYPE`         | deferred | uses `document(rels)` |
| `ELEMENT_REFERENCE`         | deferred | uses `Index-of(document(...))` |
| `CROSS_PART_COUNT`          | deferred | uses `document('Part:...')` |

So 11 of 14 typed categories are pure-XSLT-compilable; the 3
deferred categories all use `document()` traversal and stay in
the imperative evaluator (see Phase 3 for the boundary work).

1. **Codegen completes** for the 11 XSLT-compilable categories.
   One `.sch` per category remains the working organization
   (keeps compile times bounded; lets failures be localized
   during corpus regeneration).
2. **One compiled XSLT per app scope per category**: rules carry
   an `App` field (`All` / `PowerPoint` / `Word` / `Excel`); emit
   `data/openxml/compiled/<app>/<category>.xslt`. Runtime loads
   the relevant set per document type.
3. **CI step regenerates compilation on JSON change.** Adds a
   `make compile-schematron` target invoked by `make codegen`.
   Compiled XSLT is committed (it's a derived artifact, but
   diffing it under PR review catches accidental rule drift).
4. **Whole-dataset invariant** (per Spec 009): every rule of a
   compilable category in the JSON must appear in exactly one
   `.sch`. Add to `tests/test_schema_regressions.py`.

### Phase 3 — SVRL adapter, cutover, imperative-evaluator carve-out

Goal: make XSLT-first the default for compilable rules and shrink
the imperative evaluator to just the 3 cross-part categories.

1. **`semantic/svrl_adapter.py`**: parse SVRL `<failed-assert>`
   elements and emit `ValidationError` instances with the same
   `code`, `message`, `path`, `source_class`, and `severity`
   fields the imperative evaluator produces. The adapter is the
   only behavioral surface a downstream consumer can observe —
   error parity tests live here.
2. **Validator wiring**: `OpenXmlValidator._validate_semantic`
   loads the compiled XSLT bundle for the document's type, runs
   it, adapts SVRL, then runs the imperative evaluator restricted
   to the 3 cross-part categories. `--use-compiled-schematron`
   becomes the default; the legacy flag flips to
   `--use-python-schematron` as an escape hatch.
3. **Imperative evaluator carve-out**: `constraint_bridge.py`
   stops emitting `SemanticConstraint` instances for the 11
   compilable categories. The bridge keeps producing them for the
   3 cross-part categories. Update Spec 009 invariants to reflect
   the new partition.
4. **Benchmark cutover**: re-run the README benchmark and update
   the headline ratio. Capture the new warm/cold numbers in the
   release notes.

## Acceptance Criteria

### Phase 1 (spike)

1. `codegen/schematron_compiler.py` emits a valid ISO Schematron
   `.sch` for every `ATTRIBUTE_VALUE_RANGE` rule in the JSON
   corpus.
2. `scripts/compile_schematron.py` produces
   `data/openxml/compiled/range.xslt` via `lxml.isoschematron`
   without error.
3. `tests/test_schematron_compiler.py` — at minimum: emitted
   `.sch` parses, compiled XSLT loads, SVRL fires on a known
   violating fixture.
4. A side-by-side report (compiled XSLT vs Python evaluator) on
   the existing benchmark fixture shows zero missing findings and
   zero spurious findings for the spike rule type.
5. Benchmark measurement recorded in the spec or PR body.

### Phase 2 (full coverage)

1. `data/openxml/compiled/<app>/<category>.xslt` exists for each
   of the 11 compilable categories × 4 app scopes that have rules
   in that category.
2. Whole-dataset invariant test passes: every JSON rule of a
   compilable category appears in exactly one `.sch`.
3. `make compile-schematron` regenerates the compiled artifacts
   deterministically; re-running produces byte-identical output.

### Phase 3 (cutover)

1. SVRL adapter produces `ValidationError` instances
   indistinguishable (by code / message / path / source_class /
   severity) from the imperative evaluator's output for every
   compilable rule, on every fixture in `tests/`.
2. Default validation path uses the compiled XSLT; the
   `--use-python-schematron` escape hatch produces identical
   output (regression-tested against the same fixtures).
3. README benchmark shows measurable warm-latency improvement
   versus 0.7.5; numbers committed in the release notes.
4. `constraint_bridge.py` no longer emits constraints for the 11
   compilable categories; tests verify only the 3 cross-part
   categories remain.

## Out of Scope

- **Compiling the 155 schema JSON files to XSD.** Separate
  question (per ADR-004's "Out of scope"); ECMA-376 XSDs already
  exist and are known buggy. Open a follow-up spec if a specific
  speed gap motivates it.
- **WASM execution stack selection** (libxslt-WASM vs
  Saxon-JS SEF). Belongs in a downstream port's ADR/spec; this
  one only commits to "the compiled XSLT is the artifact we
  ship."
- **Compiling cross-part rules** via a custom URI resolver that
  exposes other OPC parts to the XSLT engine. Tractable but
  meaningfully more complex; deferred until the imperative
  carve-out shows it as a measurable cost.
- **Localization of SVRL messages.** SVRL `<diagnostic>` elements
  could carry localized text; the adapter emits the English `code`
  + `message` we already produce. Future work.

## Risks

- **Compilation time is "many minutes."** Mitigated by making it
  a CI step gated on JSON change, not a developer-hot-path step.
  Per-category `.sch` files keep individual compile units bounded.
- **SVRL adapter drift.** If the imperative evaluator's
  `ValidationError` shape evolves (new fields, changed codes),
  the adapter must keep up or parity regresses silently.
  Mitigated by the Phase 3 acceptance test: byte-identical
  `ValidationError` output for every fixture on both paths.
- **Compiled XSLT diffs are noisy under PR review.** A small JSON
  change can cascade into a large XSLT diff. Mitigated by
  emitting one `.xslt` per category — most JSON edits will affect
  only one or two artifacts.
- **Partition drift.** A future contributor could add a new
  `SchematronType` that needs `document()` but bucket it as
  compilable, or vice versa. Mitigated by the Phase 2 invariant
  test that asserts the exact partition (11 compilable categories,
  3 cross-part) against the live enum.
- **XSLT 1.0 vs 2.0.** The ISO Schematron skeleton lxml ships
  targets XSLT 1.0 (libxslt's level). A handful of rules might
  rely on XPath 2.0 constructs (`matches()` is fine via EXSLT;
  `distinct-values()` is also OK; harder cases would need a
  rewrite). Phase 1's spike doesn't expose this; Phase 2 may
  surface specific rules that need either rewriting or staying in
  Python. Cataloged then, not now.
