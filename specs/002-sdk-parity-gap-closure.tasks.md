# Tasks: SDK Parity Gap Closure and SDK-Free Baseline

**Spec:** [002-sdk-parity-gap-closure.md](./002-sdk-parity-gap-closure.md)

## Phase 0 (March 9-13, 2026): Measurement and Contract

- [x] Create `scripts/corpus/extract_sdk_expectations.py`
- [ ] Expand expectation extraction to at least `250` checks
- [x] Create `scripts/corpus/run_parity_snapshot.py`
- [x] Define normalized comparison tuple `(Id, ErrorType, Part, Path, DescriptionNormalized)`
- [x] Generate and commit baseline snapshot under `data/corpus/parity_baseline/v3.4.1/`
- [x] Add report summary by mismatch family (count + representative examples)

## Phase 1 (March 16-27, 2026): Core Correctness and Noise Reduction

- [ ] Add focused regression fixtures for known false-positive families
- [ ] Fix schema/content-model issues causing repeated missing `latin/ea/cs` errors
- [ ] Tighten undeclared-attribute behavior to SDK-compatible semantics
- [ ] Add nested `AlternateContent` edge-case tests
- [ ] Confirm top 3 mismatch families reduced by at least `80%`

## Phase 2 (March 30-April 10, 2026): Error Model and Semantic Tail

- [ ] Implement remaining 2 non-converted schematron edge cases
- [ ] Add stable SDK-like error ID mapping layer
- [ ] Add path normalization compatibility layer
- [ ] Align semantic error type + message templates where parity requires
- [ ] Reach at least `95%` parity on core corpus checks

## Phase 3 (April 13-17, 2026): SDK Exit for Daily CI

- [x] Add `scripts/corpus/compare_to_baseline.py`
- [x] Add PR CI gate using Python-only baseline compare
- [x] Add nightly optional SDK calibration workflow
- [ ] Document SDK as optional/manual-only in developer workflow
- [ ] Enforce no parity-drift policy in PR CI

## Operational Tasks

- [ ] Create `docs/parity_contract.md`
- [x] Pin calibration target to Open XML SDK `v3.4.1`
- [ ] Add mismatch waiver process (owner + rationale + expiry)
- [ ] Add performance regression guard for parity workloads
