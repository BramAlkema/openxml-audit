# Tasks: Whole-Dataset Invariants for SDK Bridge Fidelity

**Spec:** [009-bridge-whole-dataset-invariants.md](./009-bridge-whole-dataset-invariants.md)

## Phase 1: Enum Invariants and Bridge Downgrade Checks

- [x] Keep the live `EnumValue<T>` completeness scan in fast pytest
- [x] Add a shared helper to collect live `EnumValue<T>` uses from shipped schema JSON
- [x] Add invariant test: every live enum type resolves through `get_enum_values()`
- [x] Add invariant test: live enum attributes do not degrade to unconstrained `StringTypeValidator`
- [x] Add invariant test: wrapped enum validators still preserve enumeration semantics under `UnionTypeValidator`
- [x] Add invariant test: wrapped enum validators still preserve enumeration semantics under `VersionedTypeValidator`
- [x] Ensure invariant failures print unresolved enum type names and owning schema element/attribute

## Phase 2: Numeric and String Validator Fidelity

- [x] Add a shared helper to collect live attributes carrying `NumberValidator`
- [x] Partition numeric cases by scalar kind, min/max bounds, and `IsList`
- [x] Add invariant test: unsigned numeric branches reject negative probes
- [x] Add invariant test: decimal/float/double branches accept decimal probes
- [x] Add invariant test: list-valued numeric branches preserve list semantics
- [x] Add invariant test: numeric range bounds survive bridge conversion
- [x] Add a shared helper to collect live string-validator metadata and string-like SDK types
- [x] Add invariant test: `IsQName` metadata produces QName-constrained runtime validation
- [x] Add invariant test: NCName/ID/URI-specialized branches preserve constrained behavior
- [x] Add invariant test: pattern-only string validators do not degrade to unconstrained strings
- [x] Ensure numeric/string failures report the schema type, attribute, and probe value

## Phase 3: Versioned Validator Branch Selection

- [x] Add a scanner for attributes with version-split validator groups or validator-version metadata
- [x] Identify a stable dataset-driven sample set for Office2007, Office2010, and Office2013+ branch evolution
- [x] Add invariant test: versioned validators select the correct branch for older formats
- [x] Add invariant test: versioned validators select the correct branch for newer formats
- [x] Add invariant test: branch evolution is not flattened into a global union
- [x] Ensure failures report the attribute, active file format, and unexpected accepted/rejected value

## Phase 4: Schematron Bridge Coverage

- [x] Add a shipped-data scan over all schematron rules in `data/openxml/schematrons.json`
- [x] Assert expected bridge coverage totals from shipped schematron data
- [x] Add invariant test: forbidden fallback buckets remain at zero
- [x] Add invariant test: bridge output remains deterministic for repeated loads
- [x] Ensure failures print rule context, app scope, and rule type

## Phase 5: Ambiguous Element-Type Safety

- [x] Add a helper to enumerate ambiguous element tags from the schema registry
- [x] Select high-risk ambiguous families for probes: CustomUI, chart `ser`, chart/spreadsheet `ext`, Word `del`/`rPr`/`pPr`
- [x] Add invariant test: `get_element_constraint_for_element()` chooses the attribute-compatible candidate for CustomUI controls
- [x] Add invariant test: child-shape-sensitive candidates resolve correctly for chart/spreadsheet ambiguous tags
- [x] Add invariant test: Wordprocessing overloaded tags choose the context-appropriate candidate
- [x] Ensure failures print the tag, chosen candidate, and expected candidate set

## Phase 6: Test Layout, Performance, and Documentation

- [x] Add `tests/test_codegen_bridge_invariants.py` for the heavier whole-dataset scans
- [x] Keep lightweight loader/data invariants in `tests/test_codegen_data_resources.py`
- [x] Refactor shared scan helpers to avoid duplicated schema walks across tests
- [x] Measure runtime of the invariant suite and keep it acceptable for normal pytest runs
- [x] Split especially heavy scans if needed while keeping them in default CI coverage
- [x] Document the invariant suite as required coverage for schema/codegen changes
