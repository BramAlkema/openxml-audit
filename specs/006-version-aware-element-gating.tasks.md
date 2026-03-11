# Tasks: SDK-Version-Aware Element and Attribute Gating

**Spec:** [006-version-aware-element-gating.md](./006-version-aware-element-gating.md)

## Infrastructure

- [ ] Add `FileFormat.from_version_string(version: str) -> FileFormat | None` classmethod
- [ ] Add `FileFormat` ordinal comparison (e.g. `_OOXML_ORDER` list or `__lt__` support)
- [ ] Add `FileFormat.includes(other: FileFormat) -> bool` for version availability checks
- [ ] Unit tests for `from_version_string` mapping (all 7 versions + unknown input)
- [ ] Unit tests for ordinal comparison and `includes` logic

## Element Gating

- [ ] Verify `SchemaRegistry` exposes `introduced_version` on `ElementConstraint` from SDK JSON
- [ ] Add element version lookup in `SchemaValidator._get_validation_children()`
- [ ] Emit `Sch_UnexpectedElementContentExpectingComplex`-compatible error for version-gated elements
- [ ] Remove `_is_version_ignored_child()` method (replaced by general mechanism)
- [ ] Unit test: element with `introduced_version="Office2010"` accepted for Office2010+
- [ ] Unit test: element with `introduced_version="Office2010"` rejected for Office2007
- [ ] Unit test: element with `introduced_version=None` accepted for all versions

## Attribute Gating

- [ ] Add attribute version check in `SchemaValidator._validate_attributes()`
- [ ] Exclude version-unavailable attributes from declared set (treated as undeclared)
- [ ] Unit test: attribute with `introduced_version="Office2013"` accepted for Office2013+
- [ ] Unit test: attribute with `introduced_version="Office2013"` reported as undeclared for Office2010

## Data Audit

- [ ] Audit `introduced_version` coverage for `word/settings.xml` elements in SDK JSON
- [ ] Audit `introduced_version` coverage for `spreadsheetml` namespace elements in SDK JSON
- [ ] Audit `introduced_version` coverage for `presentationml` namespace elements in SDK JSON
- [ ] Backfill missing values from SDK `OfficeAvailabilityAttribute` if gaps found

## Validation

- [ ] Run full test suite — verify no regressions at default Office2019 format
- [ ] Run local corpus validation at Office2007 — verify `word/settings.xml` errors appear
- [ ] Run local corpus validation at Office2019 — verify no new false positives
