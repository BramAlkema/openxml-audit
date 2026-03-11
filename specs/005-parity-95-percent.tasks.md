# Tasks: Reach 95% SDK Parity

**Spec:** [005-parity-95-percent.md](./005-parity-95-percent.md)

## Phase 1: Version Gating Infrastructure

- [ ] Add `FileFormat.from_version_string()` mapping (`"Office2010"` -> `FileFormat.OFFICE_2010`, etc.)
- [ ] Add version-aware element gating in `SchemaValidator._get_validation_children()` using `ElementConstraint.introduced_version`
- [ ] Add version-aware attribute gating in `SchemaValidator._validate_attributes()` using `AttributeConstraint.introduced_version`
- [ ] Replace hardcoded `_is_version_ignored_child` with general version gating mechanism
- [ ] Add unit tests for `FileFormat.from_version_string()`

## Phase 2: Constraint Data Audit

- [ ] Verify `introduced_version` is populated in SDK JSON for `word/settings.xml` elements
- [ ] Verify `introduced_version` is populated for `spreadsheetml` elements used in `Spreadsheet.xlsx`
- [ ] Backfill missing `introduced_version` values from SDK `OfficeAvailabilityAttribute` annotations if needed
- [ ] Verify schema registry correctly loads and exposes `introduced_version` on constraints

## Phase 3: Parity Validation

- [ ] Run local parity snapshot with version gating enabled
- [ ] Verify `Document.docx` Office2007 error count moves toward 415
- [ ] Verify `Document.docx` Office2013 error count moves toward 1
- [ ] Verify `complex2010.docx` Office2007 error count moves toward 34
- [ ] Verify `Spreadsheet.xlsx` Office2007/2010/2013 error count moves toward [1, 2]
- [ ] Verify no regressions in currently-matched checks (72 checks still pass)
- [ ] Regenerate parity baseline if match rate >= 95%

## Phase 4: Cleanup

- [ ] Add targeted parity test fixtures for `TestFiles/Document.docx`
- [ ] Add targeted parity test fixtures for `TestFiles/Spreadsheet.xlsx`
- [ ] Add targeted parity test fixtures for `TestFiles/complex2010.docx`
- [ ] Waive residual mismatches (if any) with owner + rationale + expiry
- [ ] Update `specs/002-sdk-parity-gap-closure.tasks.md` follow-up items as complete
- [ ] Confirm parity gate passes on CI
