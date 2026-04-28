# Tasks: Word Compatibility Check for WML Property Element Ordering

**Spec:** [010-word-compat-element-ordering.md](./010-word-compat-element-ordering.md)

## Phase 1: CT_TrPr Subsequence Check

### Engine and module scaffolding

- [x] Create `src/openxml_audit/word/compat.py` with module docstring citing the spec and ECMA-376 as sources of truth
- [x] Define `ChildSequence` dataclass (`parent_tag`, `spec_section`, `children: tuple[str, ...]`)
- [x] Implement the subsequence-check function (single-pass cursor algorithm; returns the first out-of-order child name and the canonical position it should have appeared before, or None)
- [x] Build the constraint-table dict keyed on Clark-notation parent tag
- [x] Add `CT_TrPr` entry transcribed from ECMA-376 §17.4.79; cite the section in a comment per child group
- [x] Implement `WordCompatValidator` class wrapping the table, with a `validate(part, context)` entry point
- [x] Walk every element in the part; for each tag in the table, apply the subsequence check and emit a WARNING via `context.add_error(...)` when violated
- [x] Use the existing `ValidationErrorType.SEMANTIC` (this is not a schema violation; the schema validator passes); severity WARNING; node = the offending child local name

### Hook into validation pipeline

- [x] Identify the existing DOCX validator entry point (likely `src/openxml_audit/word/document.py`)
- [x] Determine which WordprocessingML parts to walk (main document, headers, footers, footnotes, endnotes, comments — whatever carries text content)
- [x] Call `WordCompatValidator` after schema/semantic phases, on each text-content part
- [x] Confirm no test currently asserts "exactly N errors" on a built DOCX in a way the new WARNING would break

### Tests

- [x] Create `tests/test_word_compat_ordering.py`
- [x] Unit: subsequence engine — empty observed list passes
- [x] Unit: subsequence engine — observed children in canonical order pass
- [x] Unit: subsequence engine — observed subset (skipping canonical entries) passes
- [x] Unit: subsequence engine — single reorder fails with the offending child reported
- [x] Unit: subsequence engine — repeated child stays in order, passes
- [x] Unit: subsequence engine — unknown child (not in canonical) is silently skipped
- [x] Constraint-table integrity: every entry has non-empty children, no duplicate children within an entry, every parent_tag uses `{...}local` Clark notation
- [x] Integration: build a DOCX with `cantSplit` after `tblHeader` in `trPr`; assert WARNING fires with `trPr` and `cantSplit` in the description and `§17.4.79` cited
- [x] Integration regression: `python-docx` `Document().save()` produces zero Phase 1 ordering warnings (control for false-positive detection)
- [x] Run full pytest suite; confirm green and runtime within 5% of baseline

### Documentation

- [x] Add CHANGELOG entry under `[Unreleased]` describing Phase 1 scope, severity, and the proxy-not-oracle caveat
- [x] Add a comment at the top of `compat.py` explaining the regeneration story (where in the spec to look when a new ECMA-376 revision lands)
- [x] Comment on issue #3 with the Phase 1 ship status, link to commit, and ask for the corpus to gate Phase 2

## Phase 2: CT_PPr and CT_RPr

Phase exit gate from Phase 1 must be met first: zero false positives on baseline + corpus signal (or explicit decision to ship without corpus).

- [ ] Transcribe `CT_PPr` children from ECMA-376 §17.3.1.26
- [ ] Transcribe `CT_RPr` children from ECMA-376 §17.3.2.28
- [ ] Add table entries with section citations
- [ ] Add unit tests asserting the new entries are well-formed via the integrity test
- [ ] Add at least one integration test per type — a DOCX with a known reordering that should fire WARNING
- [ ] Add a control test: representative `python-docx` content (paragraph with formatting, runs with multiple properties) produces zero new warnings
- [ ] Update CHANGELOG entry to cover Phase 2 types

## Phase 3: CT_TblPr and CT_TcPr

Phase exit gate from Phase 2 must be met first.

- [ ] Transcribe `CT_TblPr` children from ECMA-376 §17.4.60
- [ ] Transcribe `CT_TcPr` children from ECMA-376 §17.4.70
- [ ] Add table entries with section citations
- [ ] Integration tests per type covering at least one realistic reorder
- [ ] Control test: `python-docx`-built table with cell formatting produces zero new warnings
- [ ] Update CHANGELOG entry to cover Phase 3 types

## Phase 4 (Optional, gated): Empirical Word-Tolerance Calibration

Conditional on having a corpus from issue #3 follow-up. Not committed by this spec — placeholder so the work has a home if it materializes.

- [ ] Collect a corpus of "Word triggered repair" DOCX inputs with their offending property trees
- [ ] For each Phase 1–3 complex type, partition the corpus by parent type and observed deviation
- [ ] Identify systematic over-flags (XSD sequence violation but Word accepts) and under-flags (no XSD violation but Word repairs)
- [ ] Decide per type whether to extend the constraint shape (e.g., per-pair "may swap" allowances) or accept the proxy gap
- [ ] If the gap is small, document it and close the loop. If material, file a follow-up spec for an empirical oracle

## Cross-Phase Hygiene

- [x] Re-check pyright/mypy after each phase — the constraint table is plain data and shouldn't introduce typing surprises, but verify
- [x] Re-check ruff after each phase — long children tuples may need formatting attention
- [x] Per-phase: ensure CHANGELOG accurately reflects what shipped; do not let Phase 2/3 quietly tag onto the Phase 1 entry
