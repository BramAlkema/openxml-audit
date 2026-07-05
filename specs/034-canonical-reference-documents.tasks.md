# Tasks: Canonical Reference Documents — Ledger-Generated, Tier-Honest

**Spec:** [034-canonical-reference-documents.md](./034-canonical-reference-documents.md)

## Phase 1: Ledger and Emitters

- [x] Add `openxml_audit.reference.ledger` with `TIER_ORDER`, tier ranking, and `qualifies_at()`
- [x] Add `collect_ledger()` aggregating the pptx/docx/xlsx capability registries into `LedgerEntry` rows
- [x] Add `openxml_audit.reference.emitters` with `PptxSlideSource` and the four timing-feature slide bindings
- [x] Define DOCX body-block and XLSX row emitter protocols with empty Phase 1 registries
- [x] Report fade/wipe as emitter gaps (never silently dropped)

## Phase 2: Builders and Manifest

- [x] PPTX builder: repackage `timing_oracle` scaffold with selected slides renumbered behind a generated index slide
- [x] PPTX builder: rewrite `presentation.xml`, presentation rels, and `[Content_Types].xml` for the selected slide set
- [x] DOCX builder: generated body (title, per-feature sections, empty-ledger note) in a canonical package (styles/settings/fontTable/theme)
- [x] XLSX builder: index sheet (header row, feature rows, empty-ledger note) in a canonical package (shared strings/styles/theme)
- [x] Emit `<artifact>.manifest.json` with included/excluded entries and machine-readable exclusion reasons
- [x] Self-validate every built document with `OpenXmlValidator`; fail the build on any error
- [x] Keep builds byte-reproducible (no timestamps, stable ordering)

## Phase 3: CLI and Tests

- [x] `python -m openxml_audit.reference build --format {pptx,docx,xlsx,all} --minimum-tier TIER --out DIR`
- [x] `python -m openxml_audit.reference status [--json]` coverage/gap report
- [x] Tests: ledger aggregation and tier-rank qualification semantics
- [x] Tests: PPTX reference at `loadable` contains index + five probe slides and passes the validator
- [x] Tests: DOCX/XLSX references build, validate, and declare zero features honestly
- [x] Tests: manifest contents (inclusion, exclusion reasons, registered-tier fidelity)
- [x] Tests: byte-reproducibility of consecutive builds
- [x] Tests: CLI smoke (`build`, `status`, `--json`)

## Phase 4: Ladder Cycle (follow-up release)

- [ ] Run the PPTX repair oracle on the generated reference deck; commit the observation baseline
- [ ] Promote findings confirmed `preserved` to `roundtrip-preserved` with baseline provenance
- [x] Author entrance fade/wipe slides so the two animation findings gain emitters (generated at build time from the oracle-deck fragment builders; `reference/pptx_slides.py`)
- [ ] First DOCX/XLSX findings registered from oracle evidence; rebuild references
