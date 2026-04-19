# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.5.0] - 2026-04-20

### Added
- committed PPTX oracle deck scaffolds under `data/pptx_oracle/scaffolds/`,
  shipped inside the wheel so oracle starter decks can be materialized from
  packaged assets instead of scratch-only local files
- scaffold materialization helpers for committed PPTX oracle packages, with
  coverage asserting the wheel contains the expected PowerPoint package parts

### Changed
- `build_oracle_starter_deck()` and `build_timing_oracle_deck()` now rebuild
  decks from scaffold package trees instead of requiring `python-pptx` at
  runtime
- PPTX oracle builder and lab modules now target the shipped scaffold data as
  their runtime source of truth, while keeping the maintenance path available
  for scaffold regeneration
- repository linting was normalized so `ruff check .` passes again across the
  shipped package, scripts, and tests

## [0.4.3] - 2026-03-12

### Added
- pytest plugin with auto-registered fixtures (`assert_valid_pptx`,
  `assert_valid_docx`, `assert_valid_xlsx`, `assert_valid_odf`) — zero
  config, just `pip install openxml-audit` and use in tests
- GitHub Action for validating Office files in PRs (`changed-only` mode)
- Pre-commit hook for OOXML and ODF files
- Example scripts for python-pptx, openpyxl, ODF, and CI batch validation

### Changed
- 2x batch validation speedup via hot-path optimizations:
  single-candidate constraint cache, ignorable namespace passdown,
  singleton particle validators, and XML parse caching
- Warm validation of Document.docx (798K): 164ms → 101ms

## [0.4.2] - 2026-03-11

### Fixed
- 100% parity with Open XML SDK v3.4.1 (up from 98.7%) — 77/77 checks matched
  - 1 false-positive `Sem_UniqueAttributeValue` error for Document.docx at
    Office 2010 eliminated via version-aware MCE resolution in semantic validator
- Removed parity waiver (no longer needed)

## [0.4.1] - 2026-03-11

### Added
- Version-aware attribute gating: attributes introduced in later Office versions
  are flagged as undeclared when validating against earlier formats
- Unique attribute value validation (`Sem_UniqueAttributeValue`) for VML elements
- VML attribute namespace fix in schematron bridge (unnamespaced attributes
  were incorrectly resolved to the element's namespace)

### Fixed
- Parity improvement from 93.51% to 98.7% against Open XML SDK v3.4.1
  - 414 undeclared attribute errors now detected for Document.docx at Office2007
  - 33 undeclared attribute errors now detected for complex2010.docx at Office2007
  - 1 undeclared `shapeId` attribute error now detected for Spreadsheet.xlsx at Office2007
  - 1 `Sem_UniqueAttributeValue` error now detected for Document.docx at Office2007/2013

## [0.4.0] - 2026-03-11

### Added
- Version-aware element and attribute gating: elements introduced in later
  Office versions (e.g., Office 2010) are flagged when validating against
  earlier formats (e.g., Office 2007)
- Version-aware MCE resolution: `mc:Choice` branches requiring extension
  namespaces from later versions fall back to `mc:Fallback` at earlier formats
- Content model filtering by version: particle constraints from later versions
  are removed from content models during validation
- Properties validation for core (`docProps/core.xml`), extended (`docProps/app.xml`),
  and custom (`docProps/custom.xml`) package metadata
- Styles-with-effects validation for Word documents (`stylesWithEffects.xml`)

### Fixed
- False positive for `HLinks` element in extended properties validation

## [0.3.0] - 2026-03-10

### Added
- ODF semantic validation expanded from 10 to 27 rules (`ODFSEM001`--`ODFSEM027`)
- Supertheme part validation for PowerPoint (Office 2013+)
- Markup compatibility (MCE) validation phases for OOXML
- Gap-closure for ODF-OOXML validation milestones M1--M6

## [0.2.0] - 2026-03-09

### Added
- ODF (ODT/ODS/ODP) validation with staged conformance levels:
  foundation, schema-core, semantic-core, security-core
- ODF Relax NG schema-core routing with versioned schema maps
- ODF security policy validation (signatures, encryption structure)
- ODF reference calibration tooling (ODF Toolkit, OPF comparisons)
- ODF benchmarking scripts with per-phase timing breakdown
- Parity gate CI workflow enforcing SDK baseline match rate
- Parity baseline extraction, comparison, and waiver tooling
- Performance budget guard for validation timing
- Quarterly SDK update workflow for upstream tracking
- Release infrastructure (PyPI publishing, docs site)

### Changed
- Schema validation hot paths optimized with relationship caching
- Phase timing metrics added to validation pipeline

### Fixed
- Schema and relationship parity gaps aligned with SDK output
- Nested AND/OR schematron parsing

## [0.1.0] - 2026-03-08

### Added
- Initial release
- OOXML validation for PPTX, DOCX, and XLSX files
- Package structure validation (ZIP, content types, relationships)
- Schema validation (particle validators, type validators)
- Semantic validation (attribute, relationship, reference constraints)
- CLI with text, JSON, and XML output formats
- Python API with `OpenXmlValidator`, `validate_pptx`, `is_valid_pptx`
- Integration helpers: context managers, decorators, pytest fixtures
- Support for Office 2007 through Microsoft 365 format versions

[Unreleased]: https://github.com/BramAlkema/openxml-audit/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.9...v0.5.0
[0.4.3]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.2...v0.4.3
[0.4.2]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.1...v0.4.2
[0.4.1]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/BramAlkema/openxml-audit/releases/tag/v0.1.0
