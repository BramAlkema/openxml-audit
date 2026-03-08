# Spec: SDK Parity Gap Closure and SDK-Free Baseline

## Status

In Progress (March 8, 2026)

## Problem

The validator has closed many correctness bugs, but behavior is still far from Open XML SDK outcomes on known fixtures. We need a concrete path to:

1. Reach high-fidelity parity with the SDK where intended.
2. Freeze a Python-native baseline.
3. Remove SDK from daily development/CI workflows.

## Why This Matters

- Users trust outcomes that match de facto ecosystem behavior.
- Parity work has direct correctness impact (false positives/negatives).
- A frozen Python baseline enables fast, SDK-free CI after calibration.

## Current Baseline (as of March 8, 2026)

### Corpus

- SDK seed corpus catalog: `887` files listed in `data/corpus/sdk_seed/manifest.json`.
- Corpus files are materialized at runtime for calibration (for example under `/tmp/openxml-sdk-seed/files`) and are not tracked in git.
- Extensions:
  - `.docx` 419
  - `.pptx` 181
  - `.xlsx` 107
  - plus templates/macro variants.

### Expected-Outcomes Coverage

- Files with extracted high-confidence expectations: `30`
- Total expectation entries: `48`
- Current pass rate on those checks: `0/48`

### Schematron Conversion

- Parsed rules: `948`
- Bridge-converted constraints: `946/948` (`99.79%`)
- Remaining unconverted edge cases: `2`

### Known Noise Pattern

On SDK-valid files (for example `TestFiles/Plain.docx`), Python reports large schema error volume dominated by repeated missing `latin/ea/cs` child requirements in theme/font structures, indicating content model fidelity issues still exist.

## Normative References

- ECMA-376 Part 3 (MCE), 5th Edition (Dec 2015): `AlternateContent/Choice/Fallback` selection semantics.
  - https://ecma-international.org/wp-content/uploads/ECMA-376-3_5th_edition_december_2015.zip
- Open XML SDK source-of-truth implementation details.
  - https://github.com/dotnet/Open-XML-SDK
- Open XML SDK target release for this parity effort: `v3.4.1` (released January 6, 2026).
  - https://github.com/dotnet/Open-XML-SDK/releases/tag/v3.4.1
- MS-OI29500 notes on Office behavior differences.
  - https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5f0a7244-24e6-44be-b93a-502f0f061b44

## Decisions

1. Parity target is **SDK behavior**, pinned to `Open XML SDK v3.4.1` during calibration.
2. MCE branch selection must remain **Choice-first** then Fallback if no Choice qualifies.
3. Parity comparison contract is based on normalized tuple:
   - `(Id, ErrorType, Part, Path, DescriptionNormalized)`
4. After calibration, SDK becomes **optional/manual**, not required for daily CI.
5. Default strictness/profile used for parity calibration must be explicitly versioned and fixed in scripts.

## Scope

### In Scope

- Schema and semantic parity for OOXML (`docx/pptx/xlsx` + template/macro variants in corpus).
- Error-model fidelity (IDs/messages/types/paths).
- Corpus-driven gating and drift prevention.
- Performance guardrails while closing correctness gaps.

### Out of Scope

- ODF parity or OASIS ODF expansion in this spec.
- Digital signature policy/parity work.
- Feature expansion unrelated to parity closure.

## Targets

### Functional Targets

1. Expand expected-outcome coverage from `10` checks to `>= 250` checks.
2. Reach `>= 95%` parity on normalized checks in core corpus.
3. Eliminate top recurring false-positive families (theme/font child requirements and similar structural noise) on SDK-valid fixtures.
4. Close remaining 2 unconverted schematron edge cases.

### Operational Targets

1. PR CI runs SDK-free parity gate against frozen baseline (using a pinned corpus archive).
2. Nightly job runs optional SDK calibration and reports drift.
3. Regression policy: no growth in mismatch count on protected corpus tiers.

## Plan

## Phase 0 (March 9-13, 2026): Measurement and Contract

Deliverables:

- `scripts/corpus/extract_sdk_expectations.py`
  - Increase high-confidence expectation extraction from SDK tests.
- `scripts/corpus/run_parity_snapshot.py`
  - Emits normalized report with mismatch taxonomy.
- `data/corpus/parity_baseline/v3.4.1/*.json`
  - Frozen baseline snapshots.

Acceptance Criteria:

- `>= 250` expectation checks extracted.
- Stable normalized output schema documented.
- Report includes top mismatch families by count.

## Phase 1 (March 16-27, 2026): Core Correctness and Noise Reduction

Workstreams:

1. Particle/content model fidelity in schema engine.
2. Undeclared attribute parity refinement (strictness aligned with SDK behavior).
3. Theme/font and related false-positive family elimination.
4. MCE behavior validation corpus for nested AlternateContent scenarios.

Acceptance Criteria:

- Core valid fixtures (`Plain.docx`, `Presentation.pptx`, `Spreadsheet.xlsx`, plus selected strict/conformance files) have error counts within target envelope vs SDK.
- Top 3 mismatch families reduced by at least `80%`.

## Phase 2 (March 30-April 10, 2026): Error Model and Semantic Tail

Workstreams:

1. Implement remaining 2 schematron edge cases.
2. Align semantic constraint outcomes and IDs to SDK style.
3. Normalize messages and path rendering for stable diffing.

Acceptance Criteria:

- Schematron bridge conversion: `948/948`.
- `>= 95%` parity on core corpus checks.
- Mismatch report shows no high-severity unknown family.

## Phase 3 (April 13-17, 2026): SDK Exit for Daily CI

Workstreams:

1. Freeze Python baseline artifacts.
2. PR gate uses Python-only baseline compare.
3. Keep SDK compare as manual/nightly calibration path only.

Acceptance Criteria:

- SDK not required for local dev or PR CI.
- Baseline drift gate blocks regressions.
- Nightly calibration can detect and report drift deltas.

## Artifacts to Add

- `scripts/corpus/extract_sdk_expectations.py` (added)
- `scripts/corpus/run_parity_snapshot.py` (added)
- `scripts/corpus/compare_to_baseline.py` (added)
- `scripts/corpus/check_perf_budget.py` (added)
- `data/corpus/parity_baseline/v3.4.1/` (initial snapshot added)
- `data/corpus/parity_baseline/v3.4.1/perf_budget.json` (added)
- `docs/parity_contract.md` (field normalization, severity mapping, ID mapping policy) (added)
- CI workflow updates for parity gate (calibration + PR parity-gate workflows added).

## Risk Register

1. **Risk:** SDK tests do not provide enough direct expected outputs.
   - Mitigation: prioritize high-confidence tests; generate SDK snapshots for selected corpus once.
2. **Risk:** Message-level parity is brittle across SDK versions.
   - Mitigation: pin v3.4.1 during calibration; normalize descriptions; track raw + normalized.
3. **Risk:** Performance regressions while tightening schema fidelity.
   - Mitigation: benchmark gate and hot-path profiling per phase.
4. **Risk:** Office behavior diverges from SDK for edge cases.
   - Mitigation: document exception list; prefer SDK target unless explicitly overridden.

## Definition of Done

All must be true:

1. Core corpus parity `>= 95%` on normalized checks.
2. Remaining mismatch families are documented with owner + resolution/waiver rationale.
3. Baseline compare is enforced in PR CI without SDK dependency.
4. SDK calibration is optional and non-blocking for daily workflows.
