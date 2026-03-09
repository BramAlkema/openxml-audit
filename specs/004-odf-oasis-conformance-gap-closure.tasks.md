# Tasks: ODF OASIS Conformance Gap Closure

**Spec:** [004-odf-oasis-conformance-gap-closure.md](./004-odf-oasis-conformance-gap-closure.md)

## Phase 0: Corpus + External Baseline

- [ ] Expand pinned corpus in `data/odf/reference_corpus/manifest.json` with conformance edge cases
- [ ] Add fixture sets for:
  - [ ] Missing/invalid auxiliary XML members declared in manifest
  - [ ] Version-variant package examples (ODF 1.2/1.3/1.4)
  - [ ] Signature/encryption structure variants
- [ ] Wire ODF Toolkit command template in calibration environment
- [ ] Wire OPF command template in calibration environment
- [ ] Regenerate baseline run report with real external runs (no `unavailable`)
- [ ] Produce first categorized mismatch triage artifact

## Phase 1: Full Relax NG Coverage

- [ ] Add schema mapping config for ODF versions and part routing
- [ ] Implement resolver for RNG includes/imports across schema bundles
- [ ] Validate all manifest-declared XML members in schema-core mode
- [ ] Add clear diagnostics for missing/unresolvable schema routes
- [ ] Add tests for schema routing, resolver behavior, and fail/pass fixtures
- [ ] Add performance guardrails for large schema-validation runs

## Phase 2: Semantic Core Rule Registry

- [ ] Add ODF semantic rule registry module with stable rule IDs
- [ ] Implement text document core semantic checks
- [ ] Implement spreadsheet document core semantic checks
- [ ] Implement presentation document core semantic checks
- [ ] Add manifest/media-type semantic consistency checks for key parts
- [ ] Add cross-part reference integrity checks
- [ ] Add tests and fixtures for every new semantic rule family

## Phase 3: Signature + Encryption Core

- [ ] Publish `docs/odf_security_policy.md` (capabilities and non-guarantees)
- [ ] Implement signature package-structure validation
- [ ] Implement encryption package-structure validation
- [ ] Add explicit diagnostics for unsupported cryptographic states
- [ ] Add optional cryptographic verification hook path (dependency-gated)
- [ ] Add signature/encryption fixture matrix and tests

## Phase 4: Reference Calibration Gate

- [ ] Extend comparison normalization for cross-tool family grouping
- [ ] Add calibration CI workflow for ODF reference drift
- [ ] Add mismatch threshold policy and failure conditions
- [ ] Add waiver file model with expiry and owner requirements
- [ ] Generate and commit first calibration report with actionable mismatch summary

## Documentation + UX

- [ ] Add conformance-level matrix to README (`foundation` / `schema-core` / `semantic-core` / `security-core`)
- [ ] Document CLI/API switches for conformance-level selection
- [ ] Document known ODF limitations that remain after this milestone
- [ ] Add troubleshooting notes for reference-validator setup

## Exit Criteria

- [ ] Schema-core validates all in-scope manifest XML members
- [ ] Semantic-core catches representative interoperability defects for ODT/ODS/ODP
- [ ] Security-core provides structure + policy diagnostics for signature/encryption scenarios
- [ ] Reference calibration runs reproducibly with external validator data
- [ ] Drift policy is enforced by automated workflow
