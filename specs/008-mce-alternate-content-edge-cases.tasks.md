# Tasks: Nested AlternateContent Edge Cases

**Spec:** [008-mce-alternate-content-edge-cases.md](./008-mce-alternate-content-edge-cases.md)

## Phase 1: Recursive Resolution

- [ ] Update `_resolve_alternate_content` to recurse on nested `mc:AlternateContent`
- [ ] Add depth limit (max 10) to prevent infinite recursion
- [ ] Test: single-level `AlternateContent` — regression test for existing behavior
- [ ] Test: nested `AlternateContent` inside `mc:Fallback`
- [ ] Test: nested `AlternateContent` inside `mc:Choice`
- [ ] Test: three-level nesting resolves correctly
- [ ] Test: depth limit exceeded — returns empty gracefully
- [ ] Test: `Choice` elements evaluated in document order (first match wins)
- [ ] Run corpus validation — verify no regressions

## Phase 2: Version-Aware Understanding (depends on spec 006)

- [ ] Build namespace-to-version map from schema registry
- [ ] Filter `_known_namespaces` by `context.file_format` version
- [ ] Test: Office2007 validation — `Choice` requiring Office2010 namespace falls to `Fallback`
- [ ] Test: Office2010 validation — `Choice` requiring Office2010 namespace is selected
- [ ] Test: `Choice` with no `Requires` attribute is always selected
- [ ] Run corpus validation at Office2007 and Office2019 — measure error delta

## Phase 3: Edge Cases (if warranted by corpus evidence)

- [ ] Survey corpus for `mc:MustUnderstand` usage
- [ ] Survey corpus for `mc:ProcessContent` usage
- [ ] Implement `MustUnderstand` processing if corpus files exercise it
- [ ] Implement `ProcessContent` handling if corpus files exercise it
