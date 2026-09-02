# Spec: Word Compatibility Check for WML Property Element Ordering

## Status

Retired in 0.8.0 (September 2, 2026). The XSD-derived ordering proxy remains
available to research tooling, but it is no longer part of runtime validation.

## 0.8.0 Resolution

The original report supplied an ordering example but no reproducible Word
build, platform, or package after follow-up. We therefore built a Word oracle
instead of promoting the anecdote into a stronger rule. Word for Mac Microsoft
365 16.89.1 preserved every generated scenario without a repair dialog or XML
change:

| Property type | Scenarios preserved | Repaired | Dialogs |
|---|---:|---:|---:|
| `CT_TrPr` | 68/68 | 0 | 0 |
| `CT_TblPr` | 93/93 | 0 | 0 |
| `CT_TcPr` | 80/80 | 0 | 0 |
| `CT_SectPr` | 155/155 | 0 | 0 |
| **Total** | **396/396** | **0** | **0** |

That includes the issue #3 `tblHeader`/`cantSplit` order and full reversals of
each tested sequence. The runtime warning therefore asserted Word behaviour
that the project could not reproduce and created a large false-positive
surface. 0.8.0 removes it from `DocumentValidator`; the proxy table, mining
script, generators, and versioned oracle baselines are retained as evidence.

Reintroduction requires a minimal affected package plus the exact Word build
and platform, followed by a versioned oracle result that demonstrates a stable
repair boundary. XSD ordering alone is insufficient.

## Problem

This section records the original, now-disproved runtime hypothesis. It is
retained for decision history and must not be read as current validator
behaviour.

The hypothesis was that Word's "unreadable content" repair dialog fires when child elements inside certain WordprocessingML property complex types appear in an order that ECMA-376 forbids — even though the .NET Open XML SDK considers the same files valid.

The trigger is a deliberate divergence between the spec and the SDK's runtime model:

- **ISO/IEC 29500 (ECMA-376) declares** these complex types as `xs:sequence`. Children must appear in a fixed order.
- **The SDK's runtime model** treats them as `xs:all`-equivalent. It validates *which* children are allowed but not *what order* they appear in.
- **Word itself** enforces the stricter sequence interpretation. Out-of-order children trigger the silent repair dialog.

A single confirmed repro from issue #3:

```xml
<w:trPr>
  <w:tblHeader/>
  <w:cantSplit/>   <!-- must come before tblHeader per ECMA-376 §17.4.79 -->
</w:trPr>
```

Both `openxml-audit` and `DocumentFormat.OpenXml` v3.5.1 (`Microsoft365` target) report this file as valid. Word repairs it on open.

## Why This Matters

- The project's mission is roundtrip survival in the target app, not strict ECMA-376 compliance. A file that triggers Word's repair dialog is corruption from the user's perspective.
- No other known OOXML validator catches this — neither the .NET SDK nor anything that depends on it. That's a real moat for `openxml-audit` if we close the gap.
- The issue reporter has shipped a python-docx-based pipeline where this class of bug surfaces in production. Real consumer impact.
- The pattern probably extends beyond the one confirmed type. The likely affected complex types — `CT_TrPr`, `CT_PPr`, `CT_RPr`, `CT_TblPr`, `CT_TcPr` — are some of the most heavily used elements in any DOCX, so the blast radius is large.

## Normative References

- ISO/IEC 29500-1:2016 (ECMA-376 5th ed.) Part 1 — Fundamentals and Markup Language Reference, especially:
  - §17.3.1.26 — `CT_PPr` (paragraph properties)
  - §17.3.2.28 — `CT_RPr` (run properties)
  - §17.4.60 — `CT_TblPr` (table properties)
  - §17.4.70 — `CT_TcPr` (cell properties)
  - §17.4.79 — `CT_TrPr` (row properties)
- WML XSDs shipped with ECMA-376 (`wml.xsd` in the spec annexes)
- Issue #3 — original report and confirmed repro
- `DocumentFormat.OpenXml` v3.5.1 runtime model (`OpenXmlValidator` with `Microsoft365` target) — the baseline this spec needs to outperform

## Current Failure Pattern

1. The SDK's runtime model treats CT_TrPr/CT_PPr/CT_RPr/CT_TblPr/CT_TcPr children as a set, not a sequence.
2. `openxml-audit`'s schema validator consumes the SDK's model via the codegen bridge.
3. The existing `SequenceParticleValidator` in `src/openxml_audit/schema/particle.py` is fully capable of order-checking — it just isn't invoked for these types because the bridge produces a permissive constraint.
4. Out-of-order children pass through every existing check.
5. Word silently repairs the file. The user's first signal is a dialog they can't suppress.

## Decisions

1. **Add a Word-compat layer separate from schema validation.** This is empirical Office-app behavior, not OOXML schema. Mixing it into the schema validator would conflate two distinct sources of truth and make the severity story incoherent.
2. **Source the canonical orderings from ECMA-376 XSDs.** `wml.xsd` from the spec annexes is the authoritative source for `xs:sequence` definitions. Hand-transcribe per complex type with a comment citing the spec section. Do not derive from the SDK's runtime model — that's the source of the bug.
3. **Severity is WARNING, unconditional.** The SDK accepts these files and Word's tolerance is empirical, not contractual. ERROR would over-promise. WARNING flags the hazard without claiming spec violation.
4. **Encode as subsequence constraints.** Observed children must form a subsequence of the canonical order — i.e., they may skip canonical entries (every entry is optional or zero-or-more), but they may not reorder them. This matches what `xs:sequence` actually means when every particle has `minOccurs=0`, which is the case for these property types.
5. **Phase rollout, gated on the per-type signal.** Ship `CT_TrPr` first (the only type with a confirmed repro). Use it to validate the proxy-vs-oracle question before committing to the other four.
6. **Treat XSD sequence as a proxy, not an oracle.** Word's actual tolerance may be narrower (false negatives — we miss things Word repairs that conform to the XSD) or wider (false positives — we flag deviations Word handles). WARNING severity and a small empirical corpus (post-MVP) protect against both.
7. **No flag/setting for opt-in.** This is a default-on check, consistent with how #2 (stylesWithEffects bidirectional) and #4 (PPTX missing parts) shipped. A user who doesn't want Word-compat warnings can filter on severity.

## Scope

### In Scope

- `CT_TrPr`, `CT_PPr`, `CT_RPr`, `CT_TblPr`, `CT_TcPr` ordering checks against the ECMA-376 XSD sequence.
- A new `compat/` module structure for Office-app-compatibility checks that don't fit the schema/semantic split.
- A constraint table format that other complex types (and other Office apps) can extend later.
- WordprocessingML only. Hooking into the existing DOCX validation pipeline.

### Out of Scope

- PowerPoint or Excel ordering checks. The same pattern likely exists in PPTX (e.g., `CT_TextParagraphProperties`), but ship Word first and learn before generalizing.
- Occurrence count checks (`minOccurs`/`maxOccurs`). Out-of-band; would require a different signal model.
- Choice groups within sequences (e.g., `<xs:choice>` nested in `<xs:sequence>`). Most affected types don't have these; defer until we hit one that does.
- Schema-validator overhaul to honor ECMA-376 over the SDK model. That's a much larger change and not what the issue asks for.
- Empirical Word-tolerance corpus collection. Spec's design supports it; building it is a separate, post-MVP effort gated on issue #3 follow-up.

## Goals

1. Ship a check that flags the issue #3 repro (`cantSplit` after `tblHeader`) as WARNING.
2. Establish the `compat/` module pattern so future Office-app-compat checks (including the fuzzy ones in #3 and beyond) have a natural home.
3. Keep the constraint table small, hand-readable, and citable to ECMA-376 sections.
4. Lay the groundwork for corpus-driven calibration without committing to corpus collection in this spec.
5. Don't break existing fast pytest runtime — ordering checks are O(children) per element, but they run on every property element in every DOCX, so the engine has to be efficient.

## Design

### Module Layout

New module: `src/openxml_audit/word/compat.py`

Why under `word/`: this is WordprocessingML-specific. If/when we extend to PPTX, `pptx/compat.py` mirrors the pattern. A top-level `compat/` directory could come later if cross-format helpers emerge, but starting per-format avoids premature abstraction.

The validator hook point: `src/openxml_audit/word/document.py` (the existing DOCX validator entry point). The compat pass runs after schema and semantic phases on the document part.

### Constraint Representation

```python
@dataclass(frozen=True)
class ChildSequence:
    """Canonical child element ordering for a WML property complex type.

    Observed children must form a subsequence of `children` — they may skip
    entries (every WML property child is xs:minOccurs=0 in practice), but
    they may not reorder them.
    """
    parent_tag: str           # Clark-notation: "{...wordprocessingml...}trPr"
    spec_section: str         # ECMA-376 reference: "§17.4.79"
    children: tuple[str, ...] # local names in canonical order
```

The constraint table is a flat tuple of `ChildSequence` instances, one per supported complex type. The validator looks up the parent tag at runtime; types not in the table are not checked.

Phase 1 ships exactly one entry (`{...}trPr`). Phase 2 adds `pPr`/`rPr`. Phase 3 adds `tblPr`/`tcPr`. Each addition is a one-line table edit plus tests.

### Subsequence Algorithm

```
Given canonical = [c0, c1, ..., cn] and observed = [o0, o1, ..., om]:
  cursor = 0
  for o in observed:
    advance cursor through canonical until canonical[cursor] == o
    if no such position: flag o as out-of-order
    cursor += 1
```

Edge cases:
- Children not in the canonical list (extension elements, unknown elements): skip silently. The schema validator owns "is this a valid child" — compat owns ordering only.
- Repeated children (e.g., multiple `<w:cantSplit/>`): allowed by the algorithm as long as each repeat appears at or after the canonical position.
- Empty observed list: trivially passes.

### Error Reporting

Severity: WARNING.

Error message format:

> trPr child '{name}' appears at position N but ECMA-376 §17.4.79 places it before '{prior_name}' at position M — Word may flag this as unreadable content

Specific enough that the user can locate the file, the property element, and the offending child — not so verbose that it floods the output.

### Hook Into Validation Pipeline

The existing `WordValidator` (or equivalent in `word/document.py`) gets a new step:

```python
# After schema/semantic validation:
ordering_errors = WordCompatValidator().validate(part, context)
errors.extend(ordering_errors)
```

The compat validator walks the document XML once, looks up each element's tag in the constraint table, applies the subsequence check, and emits warnings.

### Performance

- Constraint table lookup: O(1) per element via dict keyed on Clark-notation tag.
- Subsequence check per matched element: O(children × canonical-length). For `trPr` the canonical length is ~10; for `pPr` it's ~30. Both negligible.
- Walk: O(elements) per document. The walk is single-pass; no XPath.

The expected overhead per DOCX is well under 1ms — far below schema validation cost, which dominates.

## Test Layout

- `tests/test_word_compat_ordering.py` — new file
  - Unit tests for the subsequence engine (positive, reorder, skip, repeat, empty, unknown-child cases)
  - Constraint-table integrity tests (every entry has non-empty `children`, no duplicates within a single sequence, every parent_tag uses `{...}local` shape)
  - Integration test using a built DOCX with the exact issue #3 repro (`cantSplit` after `tblHeader`); assert WARNING fires with `trPr` and the offending child name in the description
  - Regression test: a default `python-docx` `Document().save()` produces no Phase 1 ordering warnings
- `tests/conftest.py` — no new fixtures expected; the `_build_pkg` pattern in `test_properties.py` is sufficient for crafted DOCX inputs

## Acceptance Criteria

1. The exact repro from issue #3 (`cantSplit` after `tblHeader` inside `trPr`) reports WARNING, with the description naming the property type, the offending child, and the spec section.
2. A default `python-docx` `Document().save()` produces zero Phase 1 ordering warnings (control case for false-positive detection).
3. The `ChildSequence` table for `CT_TrPr` cites ECMA-376 §17.4.79 in a comment and lists children in the order the spec declares.
4. The compat module is importable independent of the schema validator and produces structured warnings, not bare strings.
5. Full pytest suite remains green; no existing test breaks.
6. CHANGELOG entry under Unreleased documenting the new check and its WARNING severity.

## Empirical Scan (April 28, 2026)

Mined 35 DOCX/DOTX files from the sibling [TokenMoulds](https://github.com/BramAlkema/TokenMoulds) project — a corporate template generator that emits real, Word-acceptable DOCX. 68,752 property subtrees observed. Validated each observed child ordering against the SDK Children list as a subsequence:

| Property type | Subtrees observed | Pass SDK proxy | Fail | Verdict |
|---|---:|---:|---:|---|
| `tblPr` | 4,847 | 4,847 | 0 | Proxy holds |
| `tcPr` | 28,667 | 28,667 | 0 | Proxy holds |
| `sectPr` | 40 | 40 | 0 | Proxy holds (small sample) |
| `pPr` | 11,420 | 11,395 | 25 | Proxy mostly holds (0.2% deviation) |
| `rPr` | 23,754 | 23,420 | 334 | **Proxy fails (1.4% deviation)** |
| `trPr` | 0 | 0 | 0 | No corpus data |

Phase 1 (`CT_TrPr`) produced **zero false positives** on the corpus.

The `rPr` failures cluster around `color` appearing before `sz`/`szCs`/`b`/`bCs` — a pattern that repeats across 334 observations, all from working DOCX output. Either every consumer of TokenMoulds output silently accepts Word's repair dialog (implausible given PyPI distribution and downstream test coverage) or **Word's actual `rPr` tolerance is wider than ECMA-376's `xs:sequence`**. Same conclusion for `pPr` at smaller scale.

This inverts the usual framing of issue #3. The original report ("Word stricter than SDK") is true for `trPr`. The corpus shows the opposite is also true: **Word is more permissive than the spec for `rPr` and `pPr`**. Both errors point the same direction: the XSD-derived sequence is a proxy, not an oracle.

## Rollout Plan

### Phase 1 — `CT_TrPr` (shipped, commit `5c40485`)

Goal: smallest viable thing, prove the engine + module layout, validate the proxy-vs-oracle question on a single complex type with a confirmed repro.

Outcome: issue #3 repro flagged at WARNING, zero false positives on corpus, full pytest suite green.

### Phase 2 — `CT_TblPr`, `CT_TcPr`, `CT_SectPr` (corpus-validated)

Goal: ship the types where the empirical scan shows the SDK proxy holds (100% pass on corpus). Lower risk of false positives, immediate validator coverage gain.

Phase exit gate: same shape as Phase 1 — pytest green, zero false positives on TokenMoulds corpus, integration test per type using the issue's repro shape.

### Phase 3 — Empirical-canonical mining tool + `CT_PPr`, `CT_RPr`

Goal: replace the XSD-derived proxy with a corpus-derived empirical canonical for the types where the proxy fails (`rPr` certainly, `pPr` borderline). Steps:

1. Build `scripts/mine_word_property_orderings.py` — consumes a DOCX/DOTX glob, mines property subtrees, produces an empirical canonical ordering per type, and emits a structured report.
2. Run the mining tool against TokenMoulds + any additional corpus the user can point at (Word's bundled samples, LibreOffice resaves, public DOCX archives).
3. Hand-derive a canonical ordering per type from the mined data — using the dominant patterns and validating that no large cluster contradicts the chosen order.
4. Commit the empirical canonical with the source corpus statistics in a comment for traceability.
5. Add `pPr`/`rPr` constraint entries using the empirical canonical, not the SDK Children list.

Phase exit gate: false-positive rate < 0.1% on the corpus the canonical was derived from (mostly tautological), AND the constraint catches at least one synthetic reorder that *should* trigger Word repair (validates the canonical actually constrains something).

### Phase 4 (optional) — `CT_TrPr` empirical refinement

Goal: only if Shaun (or someone) provides a corpus of "Word repaired this trPr" data points, validate or refine the Phase 1 SDK-derived `trPr` canonical. Currently Phase 1 rests on a single anecdote; a corpus would either confirm the SDK proxy or surface specific allowed swaps that the WARNING shouldn't flag.

## Risk Register

1. **Risk:** The XSD sequence is a poor proxy for Word's actual tolerance — false positives (we flag what Word accepts) or false negatives (we miss what Word repairs).
   - Mitigation: WARNING severity bounds the impact of false positives. Phase 2 gate forces a corpus check before generalizing. Definition of done explicitly does not claim oracle behavior.
2. **Risk:** A hand-transcribed sequence drifts from the spec or contains a typo.
   - Mitigation: every entry comments its ECMA-376 section. Add a "constraint-table integrity" unit test that asserts shape (non-empty, no dupes, well-formed Clark tags). Optionally cross-check against the shipped SDK schema JSON for which children are *allowed*, even though the SDK doesn't carry the order.
3. **Risk:** A new ECMA-376 revision adds a child, and the table goes stale.
   - Mitigation: the table is small enough to audit on each spec revision. Document the regeneration step in a comment at the top of the constraint module.
4. **Risk:** The compat module accumulates Word-app-compat checks beyond ordering and becomes a dumping ground.
   - Mitigation: the module is named for ordering specifically (or has a sub-namespace for ordering). Future compat checks live in sibling modules unless they share the same constraint shape.
5. **Risk:** Someone interprets WARNING as advisory and ships files that Word repairs.
   - Mitigation: error message is direct about Word triggering the repair dialog. Document the severity rationale in CHANGELOG and (eventually) in user docs.

## Open Questions

- **Choice groups inside sequences.** None of the Phase 1 types have `<xs:choice>` nested in their `<xs:sequence>`, but Phase 2/3 might. Defer until concrete; the constraint shape may need extending to `tuple[str | tuple[str, ...], ...]`.
- **Occurrence counts.** Several Phase 2/3 types have `minOccurs > 0` on some children. Pure ordering doesn't catch a missing required child. Out of scope here, but flag for a future spec.
- **Cross-part walk.** Should the compat pass run only on the main document part, or also on headers, footers, footnotes, comments, etc.? They all use the same property complex types. Recommendation: run on all WordprocessingML parts that contain text content. Validate during Phase 1 implementation.

## Definition of Done

All must be true:

1. The compat module exists at `src/openxml_audit/word/compat.py` with the subsequence engine and a `CT_TrPr` constraint entry.
2. The validation pipeline runs the compat pass on every WordprocessingML text-content part.
3. Issue #3's repro fires a WARNING; a default `python-docx` baseline does not.
4. Unit tests cover the engine (reorder, skip, repeat, empty, unknown-child).
5. The `ChildSequence` for `CT_TrPr` cites ECMA-376 §17.4.79 in source.
6. CHANGELOG documents the new check, severity, and Phase 1 scope.
7. The full pytest suite remains green and runtime overhead is below 5% on a representative DOCX corpus (the existing fixtures are sufficient as a proxy).
