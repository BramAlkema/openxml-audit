# Architecture Decision Records

This document consolidates architectural decisions for openxml-audit.

This repo owns validation, inspection, and empirical PPTX research tooling.
Converter-side emission and product decisions stay in the sibling
`svg2ooxml` repo.

---

## Mission

### Evidence-Ladder Mission (ADR-001)

`openxml-audit` answers a layered question about Office files: not just
"is this schema-valid?" but "will this survive a roundtrip in the target
app?". The ladder is the organizing principle — validation is the floor
tier; loadability, roundtrip preservation, runtime behavior, and authoring
provenance are the tiers above it. Everything in the repo exists to gather
evidence at one of those tiers, in pure Python, without depending on the
.NET SDK at runtime.

The tiers (`openxml_audit.EvidenceTier`):

1. `schema-valid` — parses against ECMA/OASIS schemas
2. `loadable` — the target app opens the file without repair
3. `roundtrip-preserved` — the app's save does not materially rewrite the intent
4. `slideshow-verified` — runtime behavior in the target app matches intent
5. `ui-authored` — the app itself produced this structure

Decision:

- tier 1 (`schema-valid`) is calibrated against authoritative external
  references (Open XML SDK, ODF specs, fixture contracts), not against any
  one generator repo
- tiers 2-5 are calibrated against curated corpora of target-app-authored
  artifacts (see ADR-002)
- the repo owns validators, package inspection, diffing, lightweight "will
  this work?" capability checks, and the evidence-gathering tools for higher
  tiers
- it does not own file-generation policy for converter products

Consequences:

- consumers can depend on `openxml-audit` without inheriting converter
  assumptions
- evidence drift is handled here, once, across all five tiers, instead of
  being reimplemented in every emitting repo
- the long-term direction is a canonical per-format reference artifact (one
  `.pptx`, one `.docx`, one `.xlsx`) aggregating every feature proven at
  tier `roundtrip-preserved` or above — no such reference document exists
  today, and producing it is the payoff of the corpus work

## PPTX Research

### PPTX Evidence: XML-First and Curated (ADR-002)

ADR-001 establishes the evidence ladder; this ADR records how the **PPTX**
layer of that ladder is organized. PPTX is the first format where empirical
evidence-gathering is needed at scale — animation/timing is the canonical
"schema-valid but PowerPoint silently rewrites it" problem — and the
conventions here are the template for future DOCX and XLSX layers.

Decision:

- durable PPTX evidence lives under `docs/pptx_oracle/`
- reusable PPTX lab tooling lives under `src/openxml_audit/pptx/` with
  entrypoints `openxml-audit-pptx-lab` and `openxml-audit-pptx-timing-oracle`
- PPTX capability findings are registered in
  `src/openxml_audit/pptx/capabilities.py` and reference `EvidenceTier` from
  the format-neutral `openxml_audit.evidence` module
- committed artifacts stay XML-first and curated; full dumps, captures, and
  scratch decks stay out of git unless explicitly promoted
- consumer repos should cite or bridge these artifacts instead of copying the
  corpus into their own docs trees

Consequences:

- PPTX evidence has one durable home tied to the shared ladder
- converter repos stay focused on emission and packaging rather than lab churn
- PPTX research tools can serve multiple consumers, not just `svg2ooxml`
- when DOCX/XLSX evidence work starts, it follows this layout: format-specific
  corpus under `docs/<format>_oracle/`, format-specific capability registry
  under `src/openxml_audit/<format>/capabilities.py`, all tiered via the same
  `EvidenceTier` enum — no new abstractions required

### Oracle Deck Scaffold Layer (ADR-003)

The earliest PPTX oracle builders used `python-pptx` directly to create a
blank deck, place a few shapes, save the package, and then patch authored
`<p:timing>` XML into the resulting slides. That was a pragmatic way to move
fast, but it blurred two different concerns:

- the evidence-bearing XML fragments we are actually trying to prove
- the disposable package scaffold used to hold those fragments

Decision:

- PPTX oracle builders should depend on a first-class scaffold layer under
  `src/openxml_audit/pptx/oracle_deck_scaffold.py`
- builders own the authored fragment semantics; the scaffold layer owns
  package persistence and slide-part patching
- `python-pptx` is treated as the current scaffold implementation, not as the
  architectural source of truth for oracle evidence
- future template-backed or fragment-backed deck scaffolds should slot in
  behind this seam without forcing every oracle builder to change again
- committed PPTX oracle decks may be materialized directly from checked-in
  scaffold package trees under `data/pptx_oracle/scaffolds/`

Consequences:

- duplicated zip-patching logic stays in one place
- the code makes the current compromise explicit instead of letting it leak
  into every builder
- committed oracle evidence can move toward checked-in templates or lower-level
  OOXML package writers without another broad refactor

## Validation Engine

### Schematron-to-XSLT Precompilation for Semantic Validation (ADR-004)

Status: Proposed — 2026-05-19 · **Speed rationale withdrawn by measurement —
2026-05-20** (see postscript). The portability rationale is untested and would
need its own ADR.

> **Postscript (2026-05-20) — the speed premise was measured and does not hold.**
>
> A spike profiled validation on real documents instead of reasoning from the
> README's warm/cold ratio. Findings:
>
> - Semantic constraint evaluation is **~13% of total runtime** at worst (≈0%
>   on many files). The **schema phase is ~76%** — on a 1.7MB authored DOCX,
>   schema 2.5s vs semantic 0.5s of a 3.3s total.
> - The typed-constraint bridge **already short-circuits to near-C speed**:
>   `AttributeMinMaxConstraint` and the other "compilable" categories run at
>   sub-microsecond per call. There is no per-rule Python *XPath* loop to
>   replace — the ADR's "per-rule Python loop is the bottleneck" step did not
>   follow from "parsing isn't the bottleneck"; "above libxml2" is dispatch
>   overhead, not constraint evaluation.
> - The one genuinely slow constraint, `UNIQUE_ATTRIBUTE`, was slow from an
>   **O(N²) algorithm**, not engine speed. XSLT 1.0 (libxslt's level) lacks
>   `distinct-values()`, so reimplementing it there would be harder, not faster.
>
> The two real levers were pure-Python and algorithmic, with byte-identical
> findings: uniqueness memoization (commit `9d8a9b0`) and multi-candidate
> constraint-resolution caching (commit `793bd50`). Combined they took the
> 1.7MB DOCX from 3.47s to 2.40s (~1.45×) — more than a perfect zero-cost XSLT
> replacement of *all* semantic eval could have achieved (≤13%).
>
> **Conclusion:** do not pursue XSLT precompilation as a performance measure.
> The compiled-XSLT *portability* argument below (WASM / Apps Script reuse) is
> logically independent and may still merit an ADR — but as a portability
> project with portability success criteria, not a speed one. The original
> proposal is preserved unedited below for the record.

The SDK semantic-rule corpus already lives in a structured form:
`data/openxml/schematrons.json` carries ~948 rules as `{Context, Test, App}`
triples — the exact shape of ISO Schematron `rule context="…"` /
`assert test="…"` — and `data/openxml/schemas/` carries 155 SDK schema JSON
files describing element and attribute particles. `codegen/schematron_loader.py`
parses each rule into one of 14 typed `SchematronType` categories, and
`codegen/schematron_bridge.py` lifts those into typed Python
`SemanticConstraint` instances which are then evaluated per-document, per-rule
in a Python loop.

The README's benchmark records 2.2× warm latency versus the .NET SDK on a 798K
DOCX. The 1.2× cold gap closes on warmth, which tells us libxml2 parsing is
not the bottleneck — the per-rule Python loop on top of it is. Microsoft's own
SDK does not validate via XSDs; it executes the same rule corpus we ship as
JSON, but inside a compiled .NET pipeline rather than an interpreted Python
one. The structural opportunity is to do what the SDK already does in spirit:
compile the rules ahead of time and run them through a fast engine at
validation time.

The natural target is ISO Schematron compiled to XSLT, not XSD. The rules are
Schematron-shaped already — value ranges, uniqueness, cross-references,
co-occurrence constraints, conditional values — and XSD cannot express most
of them. ISO Schematron has had a standard "skeleton" XSLT compiler since the
early 2000s (`github.com/Schematron/schematron`, `iso_dsdl_include.xsl` →
`iso_abstract_expand.xsl` → `iso_svrl_for_xslt1.xsl`), and lxml ships
`lxml.isoschematron` wrapping that compiler. Compilation is documented as
"many minutes for mid-sized rule sets" but the Schematron.com docs explicitly
recommend caching the result; it is a build step, not a runtime step.

Decision:

- introduce a build-time codegen stage that emits ISO Schematron `.sch` from
  `data/openxml/schematrons.json` for every rule type that is locally
  expressible (the 11 of 14 typed `SchematronType` categories that do not
  require cross-part `document()` lookups; the 3 that do — `CROSS_PART_COUNT`,
  `ELEMENT_REFERENCE`, `RELATIONSHIP_TYPE` — stay imperative)
- compile the `.sch` through the ISO Schematron skeleton pipeline to a
  validator XSLT, and ship that compiled XSLT as a package artifact under
  `data/openxml/compiled/` alongside the JSON source
- run the compiled XSLT at validation time via `lxml.isoschematron`, in a
  single C-speed pass per document, replacing the per-rule Python loop for
  the covered rule types
- keep cross-part rules in imperative Python — the existing
  `SemanticConstraint` evaluator stays the home for `CROSS_PART_COUNT`,
  `ELEMENT_REFERENCE`, `RELATIONSHIP_TYPE`, and any future rule that
  requires `document()` traversal across OOXML parts
- preserve the JSON-rules-as-source-of-truth contract; the compiled XSLT is
  a derived artifact, regenerated from JSON in CI, never hand-edited
- adapt SVRL output back to the existing error-shape so downstream evidence
  reporting does not change

Consequences:

- the per-rule Python loop stops being the dominant cost on warm validations;
  expected speedup is meaningful (closer to .NET parity, possibly past it on
  large documents)
- the compiled XSLT becomes a reusable artifact for non-Python consumers —
  any libxml2 + libxslt environment can load it, including WebAssembly
  builds. This unblocks downstream ports (Google Apps Script, browser-side
  validators) without re-porting the rule corpus; the JSON stays the single
  source of truth and the compiled XSLT is the shared execution substrate
- the codegen stage is slow but bounded; it runs only when the JSON corpus
  changes, which is when the SDK source changes
- the cross-part rule categories get a sharper boundary in code — the
  imperative evaluator no longer carries the locally-expressible rules and
  can be optimized for cross-part traversal patterns specifically
- error-message parity becomes a mechanical adapter between SVRL and the
  existing report shape, not a behavioral change; the underlying rule
  evaluation is still Microsoft's SDK corpus, only the engine is different

Out of scope for this ADR:

- compiling the 155 schema JSON files to XSD. ECMA-376 XSDs already exist
  and are known buggy; the SDK schema JSONs map more faithfully to
  imperative element/attribute checks than to XSD validation. Worth a
  separate ADR if a specific speed gap motivates it
- choosing the WASM execution stack (libxslt-WASM vs Saxon-JS SEF) for
  non-Python consumers. That belongs in the downstream port's ADR; this
  one only commits to "the compiled XSLT is the artifact we ship"
