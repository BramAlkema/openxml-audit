# Spec: Reach 95% SDK Parity

## Status

Proposed (March 11, 2026)

## Problem

The parity gate enforces zero regression but the match rate has stalled at 93.51% (72/77 checks matched). Five mismatches remain, all unwaived. Reaching the 95% target set in spec 002 requires closing at least 2 of the 5 remaining gaps.

## Why This Matters

- The 95% target was committed in spec 002 as a definition-of-done criterion.
- Each remaining mismatch represents a real correctness gap visible to users validating production files.
- Closing these gaps unblocks a v0.4.0 release with the parity milestone met.

## Current Baseline (as of March 11, 2026)

Pinned to Open XML SDK v3.4.1. Baseline snapshot: `data/corpus/parity_baseline/v3.4.1/parity_snapshot.json`.

| # | File | Kind | Version(s) | Expected | Actual | Gap |
|---|------|------|------------|----------|--------|-----|
| 1 | `TestFiles/Document.docx` | `assert_single_validator` | Office2013 | 1 | 0 | No errors emitted; SDK emits 1 |
| 2 | `TestFiles/Document.docx` | `inline_version_count` | Office2007 | 415 | 0 | No errors emitted; SDK emits 415 |
| 3 | `TestFiles/Document.docx` | `inline_version_count` | Office2013 | 1 | 0 | No errors emitted; SDK emits 1 |
| 4 | `TestFiles/Spreadsheet.xlsx` | `helper_allowed_counts` | Office2007/2010/2013 | [1, 2] | 0 | No errors emitted; SDK emits 1-2 |
| 5 | `TestFiles/complex2010.docx` | `assert_equal_validator_count` | Office2007 | 34 | 1 | 1 error emitted; SDK emits 34 |

### Mismatch Family

One active family: `Sch_UnexpectedElementContentExpectingComplex` in `/word/settings.xml` at `/settings[1]`.

### Analysis

- **Mismatches 1-3** (`Document.docx`): The SDK emits 415 errors for Office2007 and 1 for Office2013. Our validator emits 0. The dominant error family is unexpected elements in `word/settings.xml` — elements that are valid in later Office versions but not in the target version. This points directly to missing **version-aware element gating**.
- **Mismatch 4** (`Spreadsheet.xlsx`): The SDK emits 1-2 errors depending on version. Root cause likely similar — elements valid in later versions flagged as unexpected in earlier versions.
- **Mismatch 5** (`complex2010.docx`): Expected 34 errors, got 1. The single error we emit is correct but we miss 33 others. The gap is again `word/settings.xml` — many elements introduced in Office2010+ are not flagged when validating against Office2007.

### Root Cause

All 5 mismatches trace to the same root cause: **the schema validator does not gate elements or attributes by their `introduced_version`**. The `ElementConstraint` and `AttributeConstraint` dataclasses already carry an `introduced_version` field, and `FileFormat` is already threaded through `ValidationContext`, but the version check is never performed. The only version-aware logic is a single hardcoded exclusion for `w:doNotEmbedSmartTags`.

## Normative References

- ECMA-376 5th Edition: element/attribute version availability per namespace.
- Open XML SDK v3.4.1: `FileFormatVersions` enum and `OfficeAvailabilityAttribute` gating.
- Spec 002: SDK Parity Gap Closure — defines 95% target and operational framework.

## Decisions

1. Implement general version-aware element/attribute gating in `SchemaValidator`, replacing the single hardcoded `_is_version_ignored_child` override.
2. Use the existing `introduced_version` field on `ElementConstraint` and `AttributeConstraint`.
3. Map `introduced_version` strings to `FileFormat` enum values for comparison.
4. After implementation, regenerate parity baseline if match rate improves.
5. Any remaining mismatches that cannot be closed are waived with rationale.

## Scope

### In Scope

- Version-aware element gating in content model validation.
- Version-aware attribute gating in undeclared-attribute detection.
- Targeted parity fixtures for all 5 mismatched files.
- Baseline regeneration if match rate changes.

### Out of Scope

- New corpus files or SDK version upgrades.
- ODF parity work.
- Message-level parity (error descriptions/IDs).

## Plan

### Phase 1: Version Gating Infrastructure

1. Add a `FileFormat.from_version_string(version: str) -> FileFormat` mapping for `introduced_version` values (e.g. `"Office2010"` -> `FileFormat.OFFICE_2010`).
2. In `SchemaValidator._get_validation_children()`, skip child elements whose `introduced_version` is later than `context.file_format`.
3. In `SchemaValidator._validate_attributes()`, skip undeclared-attribute errors for attributes whose `introduced_version` is later than `context.file_format`.
4. Remove the hardcoded `_is_version_ignored_child` method — it becomes a special case of the general mechanism.

### Phase 2: Constraint Data Audit

1. Verify that `introduced_version` is populated in the SDK JSON constraint data for elements in `word/settings.xml`.
2. If gaps exist, backfill from SDK source metadata (`OfficeAvailabilityAttribute` annotations).
3. Verify that the schema registry loads and exposes `introduced_version` correctly.

### Phase 3: Parity Validation

1. Run parity snapshot locally against corpus with version gating enabled.
2. Verify `Document.docx` (Office2007) now emits ~415 errors.
3. Verify `complex2010.docx` (Office2007) now emits ~34 errors.
4. Verify `Spreadsheet.xlsx` (Office2007/2010/2013) now emits 1-2 errors.
5. Regenerate baseline if match rate meets 95%.

### Phase 4: Cleanup

1. Add unit tests for version gating logic.
2. Waive any residual mismatches that cannot be closed without disproportionate effort.
3. Update parity baseline snapshot.

## Acceptance Criteria

1. Match rate >= 95% (at least 74/77 checks matched).
2. Version gating is general-purpose, not hardcoded per element.
3. `_is_version_ignored_child` replaced by general mechanism.
4. All 5 mismatch files have targeted parity test fixtures.
5. Parity gate passes on CI.

## Risk Register

1. **Risk:** `introduced_version` data is incomplete or missing for key elements.
   - Mitigation: Audit SDK JSON data; backfill from SDK source annotations if needed.
2. **Risk:** Version gating introduces new false positives on valid files.
   - Mitigation: Run full corpus validation before and after; compare error counts.
3. **Risk:** Version gating changes error counts for currently-matched checks.
   - Mitigation: Phase 3 validates all 77 checks, not just the 5 mismatched ones.

## Definition of Done

1. Match rate >= 95% on normalized checks.
2. Version gating implemented as general mechanism.
3. Parity gate passes on CI with updated baseline.
4. Remaining mismatches (if any) are waived with documented rationale.
