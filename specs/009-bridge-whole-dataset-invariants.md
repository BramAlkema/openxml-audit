# Spec: Whole-Dataset Invariants for SDK Bridge Fidelity

## Status

Proposed (March 18, 2026)

## Problem

The validator has repeatedly hit the same class of correctness bug: SDK metadata is present in the shipped JSON, but the runtime bridge weakens or drops meaning while translating that metadata into Python validators.

Recent examples:

- enum short-name collisions degrading `EnumValue<T>` to plain string validation
- versioned validator branches collapsing into global unions
- `NumberValidator` type/list information being lost during conversion
- `IsQName` and similar flags being ignored

These defects are expensive to find one by one because the failure mode is usually silent fallback, not an obvious exception. A single missed mapping can turn a constrained SDK attribute into an unconstrained string while the test suite still passes.

## Why This Matters

- These bugs create false negatives: invalid documents are accepted.
- The underlying pattern is structural, not case-specific. Fixing one enum or one validator family at a time does not protect the bridge as a whole.
- The project now ships tracked SDK data and performs runtime interpretation of that data. That demands invariant tests over the full shipped dataset, not only hand-picked examples.
- Whole-dataset checks are the cheapest way to keep parity fixes from regressing silently.

## Current Failure Pattern

The bridge has four stages:

1. SDK JSON data under `data/openxml/schemas/`
2. loading/indexing in `src/openxml_audit/codegen/schema_loader.py`
3. conversion to runtime constraints in `src/openxml_audit/codegen/constraint_bridge.py`
4. enforcement in `src/openxml_audit/schema/`

The recurring failure mode is:

1. the SDK identifies a type or validator precisely
2. the runtime bridge cannot preserve that identity exactly
3. a heuristic or fallback path takes over
4. validation becomes weaker than the source metadata intended

This is a bridge-fidelity problem, not just an enum problem.

## Normative References

- Open XML SDK schema JSON data shipped in `data/openxml/schemas/`
- Open XML SDK schematron data shipped in `data/openxml/schematrons.json`
- Open XML SDK parity target pinned in `docs/parity_contract.md`
- Runtime bridge implementation:
  - `src/openxml_audit/codegen/schema_loader.py`
  - `src/openxml_audit/codegen/constraint_bridge.py`
  - `src/openxml_audit/codegen/schematron_bridge.py`
  - `src/openxml_audit/schema/types.py`
  - `src/openxml_audit/schema/validator.py`

## Decisions

1. The project will add whole-dataset invariant tests for the SDK bridge layers.
2. These tests must run entirely from tracked repo data. They must not require a live SDK checkout or network access.
3. Invariants must prefer exact identity over heuristics. Heuristics may remain as fallback, but tests must detect when they fail to preserve live metadata.
4. Silent weakening of validation is considered a test failure.
5. The invariant suite is allowed to be broader than user-visible parity tests. Its job is to protect metadata fidelity at the bridge boundary.
6. Dataset-wide assertions should be framed around live uses in shipped schemas, not theoretical schema definitions that are never referenced.

## Scope

### In Scope

- Invariant tests over shipped OOXML schema JSON and schematron data.
- Bridge correctness between SDK metadata and runtime validators.
- Detection of silent fallback from constrained types to weaker validators.
- Fast scans over the shipped dataset during pytest.

### Out of Scope

- ODF bridge invariants.
- Upstream SDK sync/generation changes.
- Full output parity against .NET SDK runtime behavior.
- Networked calibration or release workflow changes.

## Goals

1. Catch bridge regressions by scanning the full shipped dataset.
2. Prevent known failure families from reappearing under new names.
3. Keep the invariant suite fast enough for routine local and CI execution.
4. Make regressions actionable by reporting the exact schema type, attribute, or rule that failed.

## Design

### Principle: Assert on Live Metadata Paths

Each invariant should walk the shipped SDK data and assert facts about metadata that is actually consumed by runtime validation.

Examples:

- `EnumValue<T>` types that appear on real attributes
- validator declarations attached to real attributes
- element/version metadata referenced by real content models
- schematron rules that are expected to be bridge-converted

This avoids brittle “count every theoretical schema object” tests and focuses on real runtime impact.

### Invariant Family 1: Live Enum Resolution Completeness

Purpose:

- Every live `EnumValue<T>` used by the shipped schema set must resolve to enumeration values.

Assertion shape:

1. scan all shipped schema JSON files
2. collect every `DocumentFormat.OpenXml.*` enum type referenced by `EnumValue<T>` usage
3. call `get_enum_values()` for each
4. assert zero unresolved results

This directly protects against:

- short-name collisions
- missing .NET namespace to schema-prefix mappings
- accidental dependence on untracked caches such as `enums.json`

### Invariant Family 2: No Silent Downgrade of Live Enum Attributes

Purpose:

- Any live attribute declared as `EnumValue<T>` must produce a runtime validator with enumeration semantics, not a plain unconstrained string validator.

Assertion shape:

1. iterate all concrete element types and attributes from the schema registry
2. for each attribute whose SDK type contains `EnumValue<...>`
3. convert the element type through `constraint_bridge`
4. inspect the resulting `AttributeConstraint.type_validator`
5. assert that enum constraints are preserved

This should tolerate wrappers such as:

- `UnionTypeValidator`
- `VersionedTypeValidator`
- list validators containing enum members

But it must fail if the effective validator accepts arbitrary strings.

### Invariant Family 3: Live Number Validator Fidelity

Purpose:

- `NumberValidator` metadata must not lose numeric kind, sign domain, decimal behavior, or list semantics during bridging.

Assertion shape:

1. scan shipped attributes with `NumberValidator`
2. partition by:
   - declared SDK scalar type
   - validator numeric type
   - `IsList`
   - min/max facets
3. build runtime validators
4. run representative positive/negative probes derived from metadata

Representative probes should cover:

- unsigned vs signed integer
- decimal/float/double acceptance
- list-valued numeric types
- bounded ranges

This protects against “all numbers became integers” and similar bridge regressions.

### Invariant Family 4: Live String Validator Fidelity

Purpose:

- string-adjacent flags and facets such as enum, QName, token, NCName, ID, URI, and pattern semantics must survive bridge conversion.

Assertion shape:

1. scan shipped attributes with `StringValidator` or string-like SDK types
2. identify flags that require special runtime handling
3. ensure the built runtime validator exposes the expected constrained behavior

This protects against bugs like:

- `IsQName` being ignored
- token/hex-length semantics being dropped
- pattern-only validators silently degrading to unconstrained strings

### Invariant Family 5: Versioned Validator Branch Selection

Purpose:

- when SDK validators evolve by Office version, the runtime bridge must not merge incompatible branches into a version-insensitive validator.

Assertion shape:

1. scan attributes whose validators are split across versions or union groups
2. detect attributes where the effective validator depends on `context.file_format`
3. probe representative values against multiple `FileFormat` levels
4. assert branch selection changes when the metadata says it should

This protects against Office 2007/2010/2013 evolution being flattened into a global union.

### Invariant Family 6: Schematron Bridge Coverage Contract

Purpose:

- all shipped schematron rules expected to be bridge-converted must stay bridge-converted

Assertion shape:

1. load all shipped schematron rules
2. classify/bridge them
3. assert expected totals and expected zero counts for forbidden fallback buckets

This is narrower than full semantic parity; it protects conversion coverage.

### Invariant Family 7: Ambiguous Element-Type Safety

Purpose:

- ambiguous element tags must not be validated through a weaker or context-blind constraint path in places where validation has the concrete element instance available.

Assertion shape:

1. identify tags with multiple candidate SDK element types
2. build representative minimal elements that should steer selection by attributes or children
3. assert `get_element_constraint_for_element()` chooses the richer or context-appropriate candidate

This should focus on hot ambiguous families:

- custom UI controls
- chart `ser`, `ext`, `extLst`
- Wordprocessing overloaded tags such as `del`, `rPr`, `pPr`

## Test Layout

Recommended layout:

- keep lightweight data-loader invariants in `tests/test_codegen_data_resources.py`
- add a dedicated bridge invariant suite, for example:
  - `tests/test_codegen_bridge_invariants.py`
- keep validator-behavior regressions in existing focused suites when a bug needs a narrow repro

Guidelines:

- dataset scans should report the unresolved items, not just counts
- tests should prefer deterministic iteration and sorted failure output
- avoid snapshot-style golden files unless counts are intentionally part of the contract

## Performance Constraints

1. Whole-dataset invariant tests must remain fast enough for normal pytest runs.
2. Scans should reuse existing loader caches where safe.
3. Expensive probing should be limited to live referenced metadata, not every schema object.
4. If a family becomes slow, split it into a dedicated test module and document expected runtime.

## Acceptance Criteria

1. The repo contains a whole-dataset invariant suite for the OOXML bridge.
2. Live `EnumValue<T>` usage in shipped schemas resolves with zero unresolved results.
3. Live enum attributes do not silently degrade to unconstrained string validators.
4. Live numeric and string validator metadata families have bridge-level invariant coverage.
5. Version-aware validator evolution has at least one dataset-driven invariant test.
6. Schematron bridge conversion coverage is asserted from shipped data.
7. Invariant failures print the exact offending type/attribute/rule so fixes are actionable.

## Rollout Plan

### Phase 1

- keep the existing live-enum completeness test
- add enum-attribute downgrade checks
- add the first versioned-validator invariant

### Phase 2

- add numeric and string validator family scans
- add ambiguous element-type safety probes for the highest-risk tags

### Phase 3

- fold bridge invariant execution into the standard fast pytest path
- document the suite as a required regression guard for schema/codegen changes

## Risk Register

1. **Risk:** invariants overfit current schema counts rather than semantic guarantees.
   - Mitigation: assert semantic properties and unresolved-item lists, not raw totals, unless counts are intentionally contractual.
2. **Risk:** some metadata is intentionally ambiguous and cannot be resolved mechanically.
   - Mitigation: scope invariants to live consumed metadata and maintain an explicit waiver list only if needed.
3. **Risk:** broad scans become slow and discourage local execution.
   - Mitigation: use cached loaders, scan only live uses, and split heavy families into dedicated modules.
4. **Risk:** tests assert internal validator class shapes too rigidly.
   - Mitigation: assert effective behavior first, implementation type second.

## Definition of Done

All must be true:

1. The whole-dataset invariant suite exists and runs from tracked repo data only.
2. The suite catches silent weakening of live SDK metadata in the bridge layer.
3. At least enum, numeric, string, versioned-validator, and schematron conversion families are covered.
4. Failures point directly to the responsible schema type or rule.
5. The suite is fast enough to keep enabled in normal developer and CI test runs.
