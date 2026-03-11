# Spec: SDK-Version-Aware Element and Attribute Gating

## Status

Proposed (March 11, 2026)

## Problem

The Open XML SDK gates element and attribute validation by Office version — an element introduced in Office2010 is flagged as unexpected when validating against Office2007. Our schema validator ignores `introduced_version` metadata entirely, with a single hardcoded exception (`w:doNotEmbedSmartTags`). This is the root cause of all 5 remaining parity mismatches.

## Why This Matters

- Version-aware gating is the SDK's primary mechanism for version-specific validation. Without it, validating against Office2007 produces the same results as Office2019.
- Users who specify `--format Office2007` expect version-accurate results.
- This is a prerequisite for spec 005 (95% parity).

## Current State

### Infrastructure (already in place)

1. `FileFormat` enum defines Office2007 through Microsoft365 (`src/openxml_audit/errors.py`).
2. `ValidationContext.file_format` carries the target version through all validation phases.
3. `ElementConstraint.introduced_version: str | None` exists on all element constraints.
4. `AttributeConstraint.introduced_version: str | None` exists on all attribute constraints.
5. Schema registry loads constraints from SDK JSON data.

### What's Missing

1. No comparison between `introduced_version` and `context.file_format` during validation.
2. `_is_version_ignored_child()` is a hardcoded single-element override instead of a general mechanism.
3. `FileFormat` has no method to compare against version strings from constraint data.
4. `_validate_attributes()` does not check attribute version availability.

## Normative References

- Open XML SDK `OfficeAvailabilityAttribute`: decorates elements/attributes with `FileFormatVersions` flags.
- ECMA-376 namespace versioning: elements in transitional namespaces may have version-specific availability.
- SDK source: `DocumentFormat.OpenXml/Validation/Schema/SdbSchemaData.cs` — version filtering logic.

## Decisions

1. Version comparison is **ordinal**: Office2007 < Office2010 < Office2013 < ... < Microsoft365.
2. An element with `introduced_version = "Office2010"` is **valid** for Office2010 and later, **invalid** for Office2007.
3. An element or attribute with `introduced_version = None` is treated as available in **all versions** (Office2007+).
4. The error emitted for version-gated elements uses the existing `Sch_UnexpectedElementContentExpectingComplex` family to maintain parity normalization compatibility.
5. The error emitted for version-gated attributes uses the existing `Sch_UndeclaredAttribute` family.

## Design

### FileFormat Version Ordering

Add an ordinal property or comparison to `FileFormat`:

```python
class FileFormat(Enum):
    OFFICE_2007 = "office2007"
    OFFICE_2010 = "office2010"
    # ...

    @classmethod
    def from_version_string(cls, version: str) -> FileFormat | None:
        """Map SDK version strings like 'Office2010' to FileFormat."""
        mapping = {
            "Office2007": cls.OFFICE_2007,
            "Office2010": cls.OFFICE_2010,
            "Office2013": cls.OFFICE_2013,
            "Office2016": cls.OFFICE_2016,
            "Office2019": cls.OFFICE_2019,
            "Office2021": cls.OFFICE_2021,
            "Microsoft365": cls.MICROSOFT_365,
        }
        return mapping.get(version)
```

Version ordering uses `_OOXML_ORDER` list for ordinal comparison.

### Element Gating

In `SchemaValidator._get_validation_children()`, after the existing ignorable-namespace and MCE checks:

```python
# Skip elements introduced in a later Office version
if child_constraint and child_constraint.introduced_version:
    introduced = FileFormat.from_version_string(child_constraint.introduced_version)
    if introduced and not context.file_format.includes(introduced):
        # Element not available in target version — report as unexpected
        context.add_schema_error(...)
        continue
```

This replaces the hardcoded `_is_version_ignored_child()`.

### Attribute Gating

In `SchemaValidator._validate_attributes()`, when checking declared vs actual attributes:

```python
# Skip attributes not yet available in target version
if attr_constraint.introduced_version:
    introduced = FileFormat.from_version_string(attr_constraint.introduced_version)
    if introduced and not context.file_format.includes(introduced):
        continue  # Treat as undeclared for this version
```

### Constraint Data Flow

```
SDK JSON → SchemaRegistry → ElementConstraint/AttributeConstraint
                              ↓ introduced_version
SchemaValidator._get_validation_children()
                              ↓ compared to context.file_format
                         skip or validate
```

## Scope

### In Scope

- `FileFormat.from_version_string()` and ordinal comparison.
- Element version gating in content model validation.
- Attribute version gating in attribute validation.
- Removal of `_is_version_ignored_child()` hardcode.
- Unit tests for version comparison and gating logic.

### Out of Scope

- Namespace-level version gating (transitional vs strict namespaces).
- Version-aware content model changes (different particle structures per version).
- ODF version gating.
- Changes to SDK JSON generation or constraint bridge.

## Acceptance Criteria

1. Elements with `introduced_version` later than target format are flagged as unexpected.
2. Attributes with `introduced_version` later than target format are treated as undeclared.
3. Elements/attributes with `introduced_version = None` are accepted in all versions.
4. `_is_version_ignored_child()` is removed.
5. Existing tests continue to pass (default format is Office2019, so most elements are valid).
6. New unit tests cover version gating for Office2007, Office2010, Office2013.

## Risk Register

1. **Risk:** `introduced_version` data is sparsely populated in SDK JSON.
   - Mitigation: Audit coverage before implementation; backfill if needed.
2. **Risk:** Some elements are version-gated in SDK but should be accepted (SDK bug or intentional divergence).
   - Mitigation: Validate against corpus; waive specific cases if needed.
3. **Risk:** Version gating changes error output for Office2019 (default), breaking existing tests.
   - Mitigation: All currently-declared elements have `introduced_version` <= Office2019 or None; no change expected for default format.
