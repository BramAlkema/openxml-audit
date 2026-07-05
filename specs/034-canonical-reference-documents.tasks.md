# Tasks: Canonical Reference Documents — Ledger-Generated, Tier-Honest

**Spec:** [034-canonical-reference-documents.md](./034-canonical-reference-documents.md)

## Phase 1: Ledger and Emitters

- [ ] Add `openxml_audit.reference.ledger` with `TIER_ORDER`, tier ranking, and `qualifies_at()`
- [ ] Add `collect_ledger()` aggregating the pptx/docx/xlsx capability registries into `LedgerEntry` rows
- [ ] Add `openxml_audit.reference.emitters` with `PptxSlideSource` and the four timing-feature slide bindings
- [ ] Define DOCX body-block and XLSX row emitter protocols with empty Phase 1 registries
- [ ] Report fade/wipe as emitter gaps (never silently dropped)

## Phase 2: Builders and Manifest

- [ ] PPTX builder: repackage `timing_oracle` scaffold with selected slides renumbered behind a generated index slide
- [ ] PPTX builder: rewrite `presentation.xml`, presentation rels, and `[Content_Types].xml` for the selected slide set
- [ ] DOCX builder: generated body (title, per-feature sections, empty-ledger note) via `build_minimal_docx`
- [ ] XLSX builder: index sheet (header row, feature rows, empty-ledger note) via `build_minimal_xlsx`
- [ ] Emit `<artifact>.manifest.json` with included/excluded entries and machine-readable exclusion reasons
- [ ] Self-validate every built document with `OpenXmlValidator`; fail the build on any error
- [ ] Keep builds byte-reproducible (no timestamps, stable ordering)

## Phase 3: CLI and Tests

- [ ] `python -m openxml_audit.reference build --format {pptx,docx,xlsx,all} --minimum-tier TIER --out DIR`
- [ ] `python -m openxml_audit.reference status [--json]` coverage/gap report
- [ ] Tests: ledger aggregation and tier-rank qualification semantics
- [ ] Tests: PPTX reference at `loadable` contains index + five probe slides and passes the validator
- [ ] Tests: DOCX/XLSX references build, validate, and declare zero features honestly
- [ ] Tests: manifest contents (inclusion, exclusion reasons, registered-tier fidelity)
- [ ] Tests: byte-reproducibility of consecutive builds
- [ ] Tests: CLI smoke (`build`, `status`, `--json`)

## Phase 4: Ladder Cycle (follow-up release)

- [ ] Run the PPTX repair oracle on the generated reference deck; commit the observation baseline
- [ ] Promote findings confirmed `preserved` to `roundtrip-preserved` with baseline provenance
- [ ] Author entrance fade/wipe scaffold slides so the two animation findings gain emitters
- [ ] First DOCX/XLSX findings registered from oracle evidence; rebuild references
