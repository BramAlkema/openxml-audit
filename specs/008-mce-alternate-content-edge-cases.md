# Spec: Nested AlternateContent Edge Cases

## Status

Proposed (March 11, 2026)

## Problem

Markup Compatibility and Extensibility (MCE) handling in the schema validator supports basic `AlternateContent/Choice/Fallback` resolution, but does not cover edge cases defined in ECMA-376 Part 3. These edge cases include nested `AlternateContent`, multiple `Choice` elements with version-dependent `Requires`, and interactions between MCE and version-aware validation.

## Why This Matters

- MCE is how Office documents maintain backwards compatibility — newer elements are wrapped in `Choice` blocks with `Fallback` alternatives.
- Incorrect MCE resolution means we validate the wrong branch, producing false positives or missing real errors.
- The SDK's MCE handling is well-defined; our gaps are a source of subtle parity divergence.
- Spec 002 listed nested `AlternateContent` tests as pending.

## Current State

### Implemented (`src/openxml_audit/schema/validator.py`)

1. **`_resolve_alternate_content(alt)`**: Selects first qualifying `mc:Choice`, falls back to `mc:Fallback`, returns children.
2. **`_select_alternate_content_choice(alt, ns)`**: Iterates `mc:Choice` elements, checks `Requires` attribute.
3. **`_choice_is_understood(choice, prefixes)`**: Verifies all required namespace prefixes are in `_known_namespaces` (built from schema registry).
4. **`_collect_ignorable_namespaces(element)`**: Walks parent chain to collect `mc:Ignorable` prefixes.
5. **Ignorable namespace filtering**: Children in ignorable namespaces are skipped during validation.

### Known Gaps

1. **Nested `AlternateContent`**: An `AlternateContent` inside a `Choice` or `Fallback` is not recursively resolved. The current code only handles one level.
2. **Version-dependent `Requires`**: `Requires` checks use a static `_known_namespaces` set. There is no version-aware namespace availability — namespaces introduced in Office2013 should not be "understood" when validating against Office2007.
3. **Multiple `Choice` ordering**: ECMA-376 Part 3 requires `Choice` elements to be evaluated in document order. Current implementation does this correctly but has no tests verifying it.
4. **`MustUnderstand` processing**: Not implemented. Elements in `MustUnderstand` namespaces should cause a validation error if the namespace is not understood.
5. **`ProcessContent`**: Not implemented. The `mc:ProcessContent` attribute specifies elements in ignorable namespaces that should still be validated.
6. **Nested `AlternateContent` in `Fallback`**: When a `Choice` is not understood, the `Fallback` may contain its own `AlternateContent` — this must be resolved recursively.

## Normative References

- ECMA-376 Part 3, 5th Edition (December 2015): Markup Compatibility and Extensibility.
  - Section 10.2: AlternateContent processing model.
  - Section 10.2.2: Choice selection semantics.
  - Section 10.2.3: Fallback processing.
  - Section 10.3: MustUnderstand processing.
  - Section 10.4: ProcessContent processing.
- Open XML SDK: `MCSupport.cs`, `OpenXmlCompositeElement.cs` — MCE processing during validation.

## Decisions

1. Recursive `AlternateContent` resolution is the highest priority — it's the most likely gap to cause real parity divergence.
2. Version-aware namespace understanding depends on spec 006 and is deferred to Phase 2.
3. `MustUnderstand` and `ProcessContent` are lower priority — implement only if corpus evidence shows parity impact.
4. Tests use synthetic XML fixtures, not full OOXML packages, for fast iteration.

## Design

### Recursive AlternateContent Resolution

```python
def _resolve_alternate_content(self, alt: etree._Element) -> list[etree._Element]:
    ns = {"mc": MC}
    chosen = self._select_alternate_content_choice(alt, ns)
    if chosen is None:
        chosen = alt.find("mc:Fallback", ns)
    if chosen is None:
        return []
    result = []
    for child in chosen:
        if not isinstance(child.tag, str):
            continue
        if child.tag == f"{{{MC}}}AlternateContent":
            result.extend(self._resolve_alternate_content(child))  # recurse
        else:
            result.append(child)
    return result
```

### Version-Aware Namespace Understanding

After spec 006, `_known_namespaces` should be filtered by version:

```python
def _build_known_namespaces(self, file_format: FileFormat) -> set[str]:
    known = {MC}
    for ns, introduced in self._schema_registry.namespace_versions():
        if file_format.includes(introduced):
            known.add(ns)
    return known
```

This means an Office2007 validation would not "understand" Office2010 namespaces, causing `Choice` elements that require those namespaces to fall through to `Fallback`.

## Scope

### In Scope

- Recursive `AlternateContent` resolution (nested `Choice`/`Fallback`).
- Test fixtures for nested MCE scenarios.
- Version-aware namespace understanding (after spec 006).
- Document-order `Choice` evaluation tests.

### Out of Scope

- `MustUnderstand` processing (deferred unless corpus evidence).
- `ProcessContent` attribute handling (deferred unless corpus evidence).
- MCE in non-OOXML contexts (ODF does not use MCE).

## Plan

### Phase 1: Recursive Resolution

1. Update `_resolve_alternate_content` to recurse on nested `AlternateContent`.
2. Add depth limit (e.g. 10) to prevent infinite recursion on malformed input.
3. Add test fixtures:
   - Single-level `AlternateContent` (existing behavior, regression test).
   - Two-level nested `AlternateContent` in `Fallback`.
   - Two-level nested `AlternateContent` in `Choice`.
   - Three-level nesting (stress test).

### Phase 2: Version-Aware Understanding (depends on spec 006)

1. Filter `_known_namespaces` by `file_format` version.
2. Add test fixtures:
   - Office2007 validation falls through to `Fallback` when `Choice` requires Office2010 namespace.
   - Office2010 validation selects `Choice` when namespace is available.
3. Verify parity impact on corpus.

### Phase 3: Edge Cases (if warranted)

1. Assess `MustUnderstand` and `ProcessContent` impact on corpus.
2. Implement if any corpus files exercise these features.

## Acceptance Criteria

1. Nested `AlternateContent` (2+ levels) resolves correctly.
2. Depth limit prevents stack overflow on malformed input.
3. Document-order `Choice` evaluation verified by tests.
4. No regressions in existing MCE handling.
5. Version-aware namespace understanding active after spec 006.

## Risk Register

1. **Risk:** Recursive resolution changes behavior for currently-passing files.
   - Mitigation: Run corpus validation before and after; diff error counts.
2. **Risk:** Depth limit too low for legitimate deeply-nested MCE structures.
   - Mitigation: Survey corpus for maximum observed nesting depth; set limit above observed max.
3. **Risk:** Version-aware namespace filtering removes too many "understood" namespaces.
   - Mitigation: Build namespace-version map from SDK metadata; validate against known-good files per version.
