# Tasks: SDK Parity Gap Closure and SDK-Free Baseline

**Spec:** [002-sdk-parity-gap-closure.md](./002-sdk-parity-gap-closure.md)

## Phase 0 (March 9-13, 2026): Measurement and Contract

- [x] Create `scripts/corpus/extract_sdk_expectations.py`
- [x] Expand expectation extraction to at least `250` checks
- [x] Create `scripts/corpus/run_parity_snapshot.py`
- [x] Define normalized comparison tuple `(Id, ErrorType, Part, Path, DescriptionNormalized)`
- [x] Generate and commit baseline snapshot under `data/corpus/parity_baseline/v3.4.1/`
- [x] Add report summary by mismatch family (count + representative examples)

## Phase 1 (March 16-27, 2026): Core Correctness and Noise Reduction

- [x] Add focused regression fixtures for known false-positive families
- [x] Fix schema/content-model issues causing repeated missing `latin/ea/cs` errors
- [ ] Tighten undeclared-attribute behavior to SDK-compatible semantics
- [ ] Add nested `AlternateContent` edge-case tests
- [x] Confirm top 3 mismatch families reduced by at least `80%`

## Phase 2 (March 30-April 10, 2026): Error Model and Semantic Tail

- [x] Implement remaining 2 non-converted schematron edge cases
- [x] Add stable SDK-like error ID mapping layer
- [x] Add path normalization compatibility layer
- [x] Align semantic error type + message templates where parity requires
- [ ] Reach at least `95%` parity on core corpus checks

## Phase 3 (April 13-17, 2026): SDK Exit for Daily CI

- [x] Add `scripts/corpus/compare_to_baseline.py`
- [x] Add PR CI gate using Python-only baseline compare
- [x] Add nightly optional SDK calibration workflow
- [x] Document SDK as optional/manual-only in developer workflow
- [x] Enforce no parity-drift policy in PR CI

## Operational Tasks

- [x] Create `docs/parity_contract.md`
- [x] Pin calibration target to Open XML SDK `v3.4.1`
- [x] Add mismatch waiver process (owner + rationale + expiry)
- [x] Add performance regression guard for parity workloads

## Focused Follow-Up (Step 12 Base-Only Residuals)

- [x] Separate mutation-dependent expectations from base-file expectations in extraction + snapshot flows
- [ ] Add SDK-version-aware element/attribute availability gating in schema validation (Office2007/2010/2013 deltas)
- [x] Scope `w:doNotEmbedSmartTags` compatibility override by file format version (do not allow unconditionally)
- [ ] Add targeted parity fixtures/assertions for `TestFiles/Document.docx` (`Office2007`, `Office2013`)
- [ ] Add targeted parity fixtures/assertions for `TestFiles/Spreadsheet.xlsx` (`Office2007`, `Office2010`, `Office2013`)
- [ ] Add targeted parity fixtures/assertions for `TestFiles/complex0.docx` (`Office2007`, `Office2010`)
- [ ] Add targeted parity fixtures/assertions for `TestFiles/complex2010.docx` (`Office2007`, `Office2010`)
- [ ] Capture SDK-side error IDs/messages for the above files/versions during calibration and store as audit artifact
