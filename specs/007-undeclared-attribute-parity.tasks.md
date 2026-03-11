# Tasks: Undeclared-Attribute Parity Refinement

**Spec:** [007-undeclared-attribute-parity.md](./007-undeclared-attribute-parity.md)

## Phase 1: Consistent Multi-Candidate Attributes

- [ ] Compute attribute union across type candidates for multi-candidate elements
- [ ] Update `_should_validate_undeclared_attributes` to handle multi-candidate case
- [ ] Unit test: single-candidate element — undeclared attribute reported
- [ ] Unit test: multi-candidate element — attribute not in any candidate reported
- [ ] Unit test: multi-candidate element — attribute in one candidate not reported
- [ ] Run corpus validation — verify no new false positives on SDK-valid files

## Phase 2: Version-Aware Attribute Sets (depends on spec 006)

- [ ] Filter declared attribute set by `introduced_version` vs `context.file_format`
- [ ] Unit test: attribute introduced in Office2013 treated as undeclared for Office2007
- [ ] Unit test: attribute introduced in Office2013 treated as declared for Office2013+
- [ ] Run corpus validation at Office2007 — verify expected undeclared attribute errors appear

## Phase 3: Coverage Expansion

- [ ] Audit elements with empty `constraint.attributes` — identify backfill candidates
- [ ] Backfill attribute data from SDK JSON where available
- [ ] Measure corpus-wide undeclared-attribute error counts before and after expansion
