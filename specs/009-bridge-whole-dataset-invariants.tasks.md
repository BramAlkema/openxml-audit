# Tasks: Whole-Dataset Invariants for SDK Bridge Fidelity

**Spec:** [009-bridge-whole-dataset-invariants.md](./009-bridge-whole-dataset-invariants.md)

## Phase 1: Enum Invariants and Bridge Downgrade Checks

- [ ] Keep the live `EnumValue<T>` completeness scan in fast pytest
- [ ] Add a shared helper to collect live `EnumValue<T>` uses from shipped schema JSON
- [ ] Add invariant test: every live enum type resolves through `get_enum_values()`
- [ ] Add invariant test: live enum attributes do not degrade to unconstrained `StringTypeValidator`
- [ ] Add invariant test: wrapped enum validators still preserve enumeration semantics under `UnionTypeValidator`
- [ ] Add invariant test: wrapped enum validators still preserve enumeration semantics under `VersionedTypeValidator`
- [ ] Ensure invariant failures print unresolved enum type names and owning schema element/attribute

## Phase 2: Numeric and String Validator Fidelity

- [ ] Add a shared helper to collect live attributes carrying `NumberValidator`
- [ ] Partition numeric cases by scalar kind, min/max bounds, and `IsList`
- [ ] Add invariant test: unsigned numeric branches reject negative probes
- [ ] Add invariant test: decimal/float/double branches accept decimal probes
- [ ] Add invariant test: list-valued numeric branches preserve list semantics
- [ ] Add invariant test: numeric range bounds survive bridge conversion
- [ ] Add a shared helper to collect live string-validator metadata and string-like SDK types
- [ ] Add invariant test: `IsQName` metadata produces QName-constrained runtime validation
- [ ] Add invariant test: NCName/ID/URI-specialized branches preserve constrained behavior
- [ ] Add invariant test: pattern-only string validators do not degrade to unconstrained strings
- [ ] Ensure numeric/string failures report the schema type, attribute, and probe value

## Phase 3: Versioned Validator Branch Selection

- [ ] Add a scanner for attributes with version-split validator groups or validator-version metadata
- [ ] Identify a stable dataset-driven sample set for Office2007, Office2010, and Office2013+ branch evolution
- [ ] Add invariant test: versioned validators select the correct branch for older formats
- [ ] Add invariant test: versioned validators select the correct branch for newer formats
- [ ] Add invariant test: branch evolution is not flattened into a global union
- [ ] Ensure failures report the attribute, active file format, and unexpected accepted/rejected value

## Phase 4: Schematron Bridge Coverage

- [ ] Add a shipped-data scan over all schematron rules in `data/openxml/schematrons.json`
- [ ] Assert expected bridge coverage totals from shipped schematron data
- [ ] Add invariant test: forbidden fallback buckets remain at zero
- [ ] Add invariant test: bridge output remains deterministic for repeated loads
- [ ] Ensure failures print rule context, app scope, and rule type

## Phase 5: Ambiguous Element-Type Safety

- [ ] Add a helper to enumerate ambiguous element tags from the schema registry
- [ ] Select high-risk ambiguous families for probes: CustomUI, chart `ser`, chart/spreadsheet `ext`, Word `del`/`rPr`/`pPr`
- [ ] Add invariant test: `get_element_constraint_for_element()` chooses the attribute-compatible candidate for CustomUI controls
- [ ] Add invariant test: child-shape-sensitive candidates resolve correctly for chart/spreadsheet ambiguous tags
- [ ] Add invariant test: Wordprocessing overloaded tags choose the context-appropriate candidate
- [ ] Ensure failures print the tag, chosen candidate, and expected candidate set

## Phase 6: Test Layout, Performance, and Documentation

- [ ] Add `tests/test_codegen_bridge_invariants.py` for the heavier whole-dataset scans
- [ ] Keep lightweight loader/data invariants in `tests/test_codegen_data_resources.py`
- [ ] Refactor shared scan helpers to avoid duplicated schema walks across tests
- [ ] Measure runtime of the invariant suite and keep it acceptable for normal pytest runs
- [ ] Split especially heavy scans if needed while keeping them in default CI coverage
- [ ] Document the invariant suite as required coverage for schema/codegen changes
