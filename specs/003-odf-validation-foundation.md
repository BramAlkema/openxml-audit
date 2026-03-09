# Spec: ODF Validation Roadmap (Reference-Aligned, Python-Native)

## Status

Proposed (March 9, 2026)

## Problem

The repository currently advertises ODF support, but implemented checks are shallow:

1. Only `mimetype` and `META-INF/manifest.xml` presence/parse are validated.
2. No package-manifest consistency enforcement.
3. No ODF XML well-formedness sweep across declared content members.
4. No ODF-specific fixtures or test suite.
5. No conformance baseline against established ODF validators.

## Why This Matters

- Users may interpret "ODF support" as conformance-level validation.
- Corrupt or inconsistent ODF packages may pass unexpectedly.
- A reference-aligned plan avoids inventing ad-hoc ODF behavior.

## Current State (as of March 9, 2026)

- ODF formats are wired into CLI/API (`odf1.2`, `odf1.3`, extension auto-detection).
- `OdfPackage` parses manifest and reads mimetype.
- `OdfValidator` returns package structure errors only.
- No ODF-focused tests in `tests/`.

## References

### Normative

- OASIS OpenDocument v1.3 / v1.4 (namespace + package model, manifest rules).

### Reference Validators (Behavior Sources)

1. ODF Toolkit ODF Validator (Java, Apache-2.0)
2. OPF ODF Validator (Java, preservation-focused)
3. Relax NG validation using ODF schemas (for Python-native schema checks)

## Decisions

1. Keep ODF as a dedicated validator pipeline (`OdfValidator`), separate from OOXML internals.
2. Build a Python-native validator first, but validate behavior against reference tools where feasible.
3. Do not hard-depend on Java in daily CI; reference validator runs are optional calibration jobs.
4. Scope this spec to **foundation + conformance core**, not full ODF semantic parity.
5. Treat warnings as non-fatal (`is_valid` should fail only on `ERROR` severity, consistent with OOXML path).

## Scope

### In Scope

- ODF package integrity checks:
  - `mimetype` entry presence and value sanity.
  - `manifest.xml` presence and XML parseability.
  - required root manifest entry and media-type consistency.
  - duplicate `manifest:file-entry` `full-path` detection.
  - manifest references to missing package members.
  - package content members missing from manifest (with defined exclusions).
- ODF XML parse checks:
  - core members (`content.xml`, `styles.xml`, `meta.xml`, `settings.xml`) if present.
  - manifest-declared XML members.
- Optional Relax NG validation hooks for ODF content members.
- ODF fixtures and tests.
- Optional reference comparison harness against ODF Toolkit/OPF for selected corpus.

### Out of Scope

- Full ODF semantic/business-rule validation parity.
- Digital signature cryptographic validation.
- Encrypted-package policy beyond clear diagnostics.
- Perfect message-level parity with external validators.

## Validation Levels

The implementation will expose explicit depth levels:

1. `foundation` (default for this milestone): package + manifest + XML well-formedness.
2. `schema` (optional): Relax NG validation for declared XML members.
3. `reference-compare` (offline/calibration): compare outcomes with ODF Toolkit/OPF.

## Functional Targets

1. Detect and classify package/manifest integrity failures deterministically.
2. Parse-check all in-scope XML members and attribute errors to concrete `part_uri`.
3. Add ODF test fixtures for valid and representative invalid cases.
4. Add optional reference-comparison tooling for drift tracking.
5. Ensure ODF `is_valid` semantics match OOXML validator semantics.

## Plan

## Phase 0 (March 10-11, 2026): Fixtures + Test Harness

Deliverables:

- `tests/fixtures/odf/` valid and invalid archives.
- New tests:
  - `tests/test_odf_package.py`
  - `tests/test_odf_validator.py`
  - CLI auto-routing tests for ODF extensions.

Acceptance Criteria:

- At least one valid fixture each for `.odt`, `.ods`, `.odp`.
- Invalid fixtures for missing/invalid `mimetype`, missing/malformed `manifest.xml`, and manifest/package mismatch.
- New tests fail before implementation and pass after.

## Phase 1 (March 12-14, 2026): Foundation Integrity Engine

Workstreams:

1. Manifest root-entry + mimetype consistency checks.
2. Duplicate file-entry checks.
3. Manifest-to-zip and zip-to-manifest consistency checks.
4. Severity policy for strict/permissive execution.

Acceptance Criteria:

- Invalid fixtures produce stable, actionable `PACKAGE` errors.
- Valid fixtures pass with zero `ERROR` outcomes.
- `is_valid` only fails on `ERROR` severity.

## Phase 2 (March 17-19, 2026): XML Parse Sweep + Optional RNG Hook

Workstreams:

1. Parse-check core and manifest-declared XML members.
2. Add optional Relax NG validation pathway (feature-gated, no hard CI dependency).
3. Improve `part_uri` attribution and deduplicate repeated parse noise.

Acceptance Criteria:

- Broken XML members are reported as `SCHEMA` errors with correct `part_uri`.
- No duplicate error spam for a single malformed member.
- Relax NG path can be toggled on/off and is tested where dependencies allow.

## Phase 3 (March 20-22, 2026): Reference Alignment and Documentation

Workstreams:

1. Add script(s) to run optional ODF Toolkit/OPF comparisons on selected corpus.
2. Produce mismatch taxonomy for Python vs reference tools.
3. Update README to describe actual ODF depth and limitations.

Acceptance Criteria:

- Reference comparison report is reproducible for a pinned sample corpus.
- README language no longer overstates conformance depth.
- ODF tests run in default `pytest` and remain stable.

## Artifacts to Add or Update

- `src/openxml_audit/odf/package.py`
- `src/openxml_audit/odf/validator.py`
- `tests/fixtures/odf/*`
- `tests/test_odf_package.py`
- `tests/test_odf_validator.py`
- Optional tooling:
  - `scripts/odf/run_reference_validators.py`
  - `scripts/odf/compare_reference_results.py`
  - `docs/odf_validation_contract.md`
- `README.md` (ODF capability + limitations)

## Risks

1. **Risk:** External validator output formats are unstable across versions.
   - Mitigation: pin reference validator versions; normalize outputs.
2. **Risk:** Relax NG support increases dependency complexity.
   - Mitigation: keep RNG validation optional and feature-gated.
3. **Risk:** Overfitting to one reference tool behavior.
   - Mitigation: compare against at least two reference sources where possible.

## Definition of Done

All must be true:

1. Foundation-level ODF package + XML integrity checks are implemented and tested.
2. ODF `is_valid` semantics are consistent with OOXML path (`ERROR`-only fatal).
3. README accurately describes implemented ODF depth.
4. Optional reference-comparison workflow exists and produces a reproducible report.
5. Follow-up gaps (deep schema semantics/parity) are explicitly tracked.
