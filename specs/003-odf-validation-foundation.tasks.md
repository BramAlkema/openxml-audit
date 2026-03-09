# Tasks: ODF Validation Roadmap (Reference-Aligned, Python-Native)

**Spec:** [003-odf-validation-foundation.md](./003-odf-validation-foundation.md)

## Phase 0: Fixtures + Test Harness

- [x] Add valid fixtures:
  - [x] `tests/fixtures/odf/valid/minimal.odt`
  - [x] `tests/fixtures/odf/valid/minimal.ods`
  - [x] `tests/fixtures/odf/valid/minimal.odp`
- [x] Add invalid fixtures:
  - [x] Missing `mimetype`
  - [x] Invalid `mimetype` value
  - [x] Missing `META-INF/manifest.xml`
  - [x] Malformed `manifest.xml`
  - [x] Manifest entry points to missing part
  - [x] Package contains XML part not declared in manifest
  - [x] Duplicate `manifest:file-entry` for same `full-path`
  - [x] Root manifest entry missing
  - [x] Root media type mismatches `mimetype`
  - [x] Broken `content.xml`
- [x] Create `tests/test_odf_package.py`
- [x] Create `tests/test_odf_validator.py`
- [x] Add CLI auto-detection tests for ODF extensions and validator routing

## Phase 1: Foundation Integrity Engine

- [x] Implement root manifest entry presence check (`full-path="/"`)
- [x] Implement root media-type <-> `mimetype` consistency check
- [x] Detect duplicate `manifest:file-entry` `full-path` values
- [x] Validate manifest-referenced members exist in ZIP
- [x] Validate package members are declared in manifest (with explicit exclusions)
- [x] Add strict/permissive severity mapping tests
- [x] Align ODF `ValidationResult.is_valid` behavior to ERROR-only fatal

## Phase 2: XML Parse Sweep + Optional RNG Hook

- [x] Parse-check core members (`content.xml`, `styles.xml`, `meta.xml`, `settings.xml`) when present
- [x] Parse-check manifest-declared XML members
- [x] Deduplicate repeated parse errors for same member
- [x] Ensure `part_uri` attribution for XML parse errors
- [x] Add feature-gated Relax NG validation path
- [x] Add tests for Relax NG path (or dependency-gated skip tests)

## Phase 3: Reference Alignment + Documentation

- [x] Add `scripts/odf/run_reference_validators.py` (ODF Toolkit / OPF runners)
- [x] Add `scripts/odf/compare_reference_results.py` (normalized diff)
- [x] Add `docs/odf_validation_contract.md` (normalization and severity contract)
- [x] Add a pinned sample corpus for reference comparison
- [x] Generate first baseline mismatch report and commit summary artifact
- [x] Update README ODF section to reflect real validation depth
- [x] Document explicit ODF limitations and follow-up scope

## Post-Milestone Follow-Up

- [ ] Expand Relax NG coverage beyond core members
- [ ] Define ODF semantic rule roadmap (cross-part/content constraints)
- [ ] Define signature and encryption handling policy
