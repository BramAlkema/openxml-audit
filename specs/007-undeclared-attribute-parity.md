# Spec: Undeclared-Attribute Parity Refinement

## Status

Proposed (March 11, 2026)

## Problem

The undeclared-attribute detection in `SchemaValidator` is deliberately conservative — it only reports undeclared attributes when an element has a single unambiguous type candidate in the schema registry. This avoids false positives from incomplete SDK metadata but also means we miss legitimate undeclared-attribute errors that the SDK catches.

Additionally, undeclared-attribute detection does not account for attribute version availability (see spec 006), so attributes introduced in later Office versions are not flagged when validating against earlier versions.

## Why This Matters

- Undeclared attributes are a common source of interoperability failures in generated Office files.
- Conservative detection means users get no warning about attributes that will cause issues.
- The SDK consistently flags undeclared attributes; our silence is a parity gap.

## Current State

### Implementation (`src/openxml_audit/schema/validator.py`)

```python
def _should_validate_undeclared_attributes(self, element, constraint):
    if not constraint.attributes:
        return False  # Skip elements with no declared attributes
    tag = element.tag
    if tag not in self._undeclared_validation_cache:
        candidates = self._schema_registry.get_element_type_candidates(tag)
        self._undeclared_validation_cache[tag] = len(candidates) == 1
    return self._undeclared_validation_cache[tag]
```

**Filters applied during attribute validation:**
- Skip attributes in `mc:` (markup compatibility) namespace.
- Skip attributes in `xml:` namespace.
- Skip attributes in foreign namespaces (namespace differs from element's namespace).

**Known conservative behavior:**
- Elements with multiple type candidates are skipped entirely.
- Elements with no declared attributes are skipped entirely.
- No version-aware filtering of declared attribute sets.

### Error Mapping

Errors use description pattern `"The '{attr}' attribute is not declared."` which maps to `Sch_UndeclaredAttribute` in parity normalization.

## Normative References

- Open XML SDK: `SchemaValidator.ValidateAttributes()` — validates against version-filtered declared attributes.
- Spec 006: Version-Aware Element and Attribute Gating — prerequisite for version-aware undeclared detection.

## Decisions

1. Tighten undeclared-attribute detection incrementally — start with elements that have reliable metadata.
2. Depend on spec 006 for version-aware attribute filtering.
3. Do not remove the single-candidate guard entirely — extend it to cover multi-candidate elements where the attribute sets are consistent across candidates.
4. Track false-positive rate on corpus to prevent regressions.

## Design

### Phase 1: Consistent Multi-Candidate Attributes

For elements with multiple type candidates, compute the **intersection** of declared attributes across all candidates. Attributes not in the intersection for any candidate are safely undeclared:

```python
def _should_validate_undeclared_attributes(self, element, constraint):
    if not constraint.attributes:
        return False
    tag = element.tag
    if tag not in self._undeclared_validation_cache:
        candidates = self._schema_registry.get_element_type_candidates(tag)
        if len(candidates) == 1:
            self._undeclared_validation_cache[tag] = True
        elif len(candidates) > 1:
            # Union of all declared attributes across candidates
            all_declared = set()
            for c in candidates:
                all_declared.update(a.qualified_name for a in c.attributes)
            # Safe to report undeclared if attribute not in ANY candidate
            self._undeclared_validation_cache[tag] = ("multi", all_declared)
        else:
            self._undeclared_validation_cache[tag] = False
    return self._undeclared_validation_cache[tag]
```

### Phase 2: Version-Aware Attribute Sets

After spec 006 is implemented, filter the declared attribute set by version before checking for undeclared attributes. An attribute with `introduced_version="Office2013"` is not in the declared set when validating against Office2007.

### Phase 3: Coverage Expansion

Review elements currently skipped due to empty `constraint.attributes` and determine if SDK JSON data can be backfilled.

## Scope

### In Scope

- Multi-candidate attribute union for safer undeclared detection.
- Integration with version-aware attribute gating (spec 006).
- Corpus validation to measure false-positive impact.

### Out of Scope

- Attribute type validation changes (value patterns, ranges).
- Required-attribute detection improvements.
- Changes to the attribute constraint data model.

## Acceptance Criteria

1. Undeclared-attribute detection covers multi-candidate elements where attribute sets are consistent.
2. No new false positives on SDK-valid corpus files.
3. Version-aware attribute filtering active when spec 006 is implemented.
4. Parity normalization for `Sch_UndeclaredAttribute` continues to work.

## Risk Register

1. **Risk:** Multi-candidate union approach flags attributes that are valid in one candidate type.
   - Mitigation: Use union (not intersection) of declared sets — only flag attributes not in ANY candidate.
2. **Risk:** SDK JSON attribute data is incomplete for some elements.
   - Mitigation: Maintain conservative fallback; only expand where data is reliable.
