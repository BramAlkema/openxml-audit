# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/BramAlkema/openxml-audit/releases/tag/v0.1.0
