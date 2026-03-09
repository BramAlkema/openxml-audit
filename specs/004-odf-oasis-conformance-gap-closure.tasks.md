# Tasks: ODF OASIS Conformance Gap Closure

**Spec:** [004-odf-oasis-conformance-gap-closure.md](./004-odf-oasis-conformance-gap-closure.md)

## Phase 0: Corpus + External Baseline

- [x] Expand pinned corpus in `data/odf/reference_corpus/manifest.json` with conformance edge cases
- [x] Add fixture sets for:
  - [x] Missing/invalid auxiliary XML members declared in manifest
  - [x] Version-variant package examples (ODF 1.2/1.3/1.4)
  - [x] Signature/encryption structure variants
- [x] Wire ODF Toolkit command template in calibration environment
- [x] Wire OPF command template in calibration environment
- [x] Regenerate baseline run report with real external runs (no `unavailable`)
- [x] Produce first categorized mismatch triage artifact

## Phase 1: Full Relax NG Coverage

- [x] Add schema mapping config for ODF versions and part routing
- [x] Implement resolver for RNG includes/imports across schema bundles
- [x] Validate all manifest-declared XML members in schema-core mode
- [x] Add clear diagnostics for missing/unresolvable schema routes
- [x] Add tests for schema routing, resolver behavior, and fail/pass fixtures
- [x] Add performance guardrails for large schema-validation runs

## Phase 2: Semantic Core Rule Registry

- [x] Add ODF semantic rule registry module with stable rule IDs
- [x] Implement text document core semantic checks
- [x] Implement spreadsheet document core semantic checks
- [x] Implement presentation document core semantic checks
- [x] Add manifest/media-type semantic consistency checks for key parts
- [x] Add cross-part reference integrity checks
- [x] Add tests and fixtures for every new semantic rule family

## Phase 3: Signature + Encryption Core

- [x] Publish `docs/odf_security_policy.md` (capabilities and non-guarantees)
- [x] Implement signature package-structure validation
- [x] Implement encryption package-structure validation
- [x] Add explicit diagnostics for unsupported cryptographic states
- [x] Add optional cryptographic verification hook path (dependency-gated)
- [x] Add signature/encryption fixture matrix and tests

## Phase 4: Reference Calibration Gate

- [x] Extend comparison normalization for cross-tool family grouping
- [x] Add calibration CI workflow for ODF reference drift
- [x] Add mismatch threshold policy and failure conditions
- [x] Add waiver file model with expiry and owner requirements
- [x] Generate and commit first calibration report with actionable mismatch summary

## Documentation + UX

- [x] Add conformance-level matrix to README (`foundation` / `schema-core` / `semantic-core` / `security-core`)
- [x] Document CLI/API switches for conformance-level selection
- [x] Document known ODF limitations that remain after this milestone
- [x] Add troubleshooting notes for reference-validator setup

## Exit Criteria

- [x] Schema-core validates all in-scope manifest XML members
- [x] Semantic-core catches representative interoperability defects for ODT/ODS/ODP
- [x] Security-core provides structure + policy diagnostics for signature/encryption scenarios
- [x] Reference calibration runs reproducibly with external validator data
- [x] Drift policy is enforced by automated workflow
