<p align="center">
  <img src="docs/logo.svg" alt="OpenXML Audit" width="96">
</p>

<h1 align="center">OpenXML Audit</h1>

[![PyPI](https://img.shields.io/pypi/v/openxml-audit)](https://pypi.org/project/openxml-audit/)
[![Downloads](https://img.shields.io/pypi/dm/openxml-audit)](https://pypi.org/project/openxml-audit/)
[![Python](https://img.shields.io/pypi/pyversions/openxml-audit)](https://pypi.org/project/openxml-audit/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![CI](https://github.com/BramAlkema/openxml-audit/actions/workflows/parity-gate.yml/badge.svg)](https://github.com/BramAlkema/openxml-audit/actions/workflows/parity-gate.yml)
[![SDK Parity](https://img.shields.io/badge/SDK%20parity-100%25-brightgreen)](docs/parity_contract.md)
[![pytest](https://img.shields.io/badge/pytest-plugin-orange)](https://pypi.org/project/openxml-audit/)

Validate OOXML (PPTX/DOCX/XLSX) and ODF files in pure Python — no .NET required.

A Python port of Microsoft's [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) validation logic. Check whether generated or modified Office files will open cleanly, directly from Python scripts, CI pipelines, or anywhere .NET isn't practical.

Also supports OASIS OpenDocument Format (ODT/ODS/ODP) with staged conformance levels.

## Features

- **OOXML Validation**: Package structure, schema, semantic, properties, and format-specific checks for PPTX/DOCX/XLSX — 100% parity with Open XML SDK v3.4.1 without the .NET dependency
- **ODF Validation**: Staged conformance levels — foundation, schema-core (Relax NG), semantic-core, and security-core for ODT/ODS/ODP
- **Multiple Output Formats**: Text, JSON, and XML output
- **Performance Tooling**: Per-phase timing breakdown, benchmark scripts for both OOXML and ODF
- **Flexible Integration**: Context managers, decorators, and pytest fixtures

## Why validate?

Libraries that generate Office files routinely produce corrupt output — python-pptx has 12+ open corruption issues, docxtpl has 7, XlsxWriter 25+. These surface as "PowerPoint found a problem" dialogs for end users or silent failures in CI. With AI agents now generating slides and reports, the problem is getting worse.

openxml-audit catches these before your users do — same checks Microsoft's SDK runs, in pure Python.

| Ecosystem | Examples | How openxml-audit helps |
|-----------|----------|------------------------|
| File generators | python-pptx, python-docx, openpyxl, XlsxWriter | Validate output in tests and CI — catch corruption before release |
| Template engines | docxtpl, pptx-template | Jinja2 rendering can break XML structure — validate after render |
| Data pipelines | pandas `to_excel`, tablib, django-import-export | Assert valid exports in pipeline tests |
| AI/LLM agents | Auto-PPT, GenFilesMCP, Docling | AI-generated Office files are unreliable — validate and retry |
| Government / ODF | Suite Numerique, odfpy | ODF conformance for EU regulatory requirements |

## Installation

```bash
pip install openxml-audit
```

Or install from source:

```bash
git clone https://github.com/BramAlkema/openxml-audit.git
cd openxml-audit
pip install -e .
```

## Quick Start

### Command Line

```bash
# Validate a single file
openxml-audit presentation.pptx

# Validate an OASIS OpenDocument file
openxml-audit document.odt

# Validate with JSON output
openxml-audit presentation.pptx --output json

# Validate with XML output
openxml-audit presentation.pptx --output xml

# Validate all matching files in a directory
openxml-audit ./presentations/ --recursive

# Validate against a specific Office version
openxml-audit presentation.pptx --format Office2007

# Limit maximum errors reported
openxml-audit presentation.pptx --max-errors 10
```

### Python API

```python
from openxml_audit import validate_pptx, is_valid_pptx, OpenXmlValidator

# Quick check
if is_valid_pptx("presentation.pptx"):
    print("File is valid!")

# Detailed validation
result = validate_pptx("presentation.pptx")
if not result.is_valid:
    print(f"Found {result.error_count} errors, {result.warning_count} warnings")
    for error in result.errors:
        print(f"  [{error.severity.value}] {error.description}")

# With custom options
from openxml_audit import FileFormat

validator = OpenXmlValidator(
    file_format=FileFormat.OFFICE_2019,
    max_errors=100,
    schema_validation=True,
    semantic_validation=True,
)
result = validator.validate("presentation.pptx")
```

## ODF Validation Depth

ODF validation is staged by explicit conformance level.

| Level | Includes | Does not include |
|---|---|---|
| `foundation` | package/manifest integrity + XML parse sweep | Relax NG schema-core routing, semantic-core rules, security-core checks |
| `schema-core` | foundation + Relax NG validation for routed XML members | semantic-core and security-core checks |
| `semantic-core` | foundation + semantic-core rule families (`ODFSEM*`) | Relax NG schema-core routing, security-core checks |
| `security-core` | semantic-core + signature/encryption structural checks (`ODFSEC*`) | full cryptographic trust guarantees unless crypto verification backend is configured |

Rule registry and policy references:

- semantic rule IDs: `openxml_audit.odf.get_odf_semantic_rules()`
- security policy: `docs/odf_security_policy.md`
- reference calibration/drift contract: `docs/odf_validation_contract.md`

### CLI Conformance Selection

Use `--odf-level` when validating ODF files:

```bash
# foundation
openxml-audit file.odt --validator odf --odf-level foundation

# semantic-core (default)
openxml-audit file.odt --validator odf --odf-level semantic-core

# security-core
openxml-audit file.odt --validator odf --odf-level security-core
```

Schema-core requires a schema-route JSON file:

```bash
openxml-audit file.odt \
  --validator odf \
  --odf-level schema-core \
  --odf-schema-routes schemas/odf/routes.json
```

`--odf-schema-routes` accepts either shape:

- versioned mapping:
  - `{"1.3": {"content.xml": "schemas/odf/1.3/content.rng"}}`
- flat legacy mapping:
  - `{"content.xml": "schemas/odf/content.rng"}`

Security-core crypto verification hook:

```bash
openxml-audit file.odt \
  --validator odf \
  --odf-level security-core \
  --odf-verify-cryptography
```

### API Conformance Selection

```python
from openxml_audit import FileFormat
from openxml_audit.odf import OdfValidator

# foundation
foundation = OdfValidator(
    file_format=FileFormat.ODF_1_3,
    schema_validation=False,
    semantic_validation=False,
    security_validation=False,
)

# schema-core (routes required)
schema_core = OdfValidator(
    file_format=FileFormat.ODF_1_3,
    schema_validation=True,
    semantic_validation=False,
    security_validation=False,
    relaxng_validation=True,
    schema_routes={"1.3": {"content.xml": "schemas/odf/1.3/content.rng"}},
)

# semantic-core
semantic_core = OdfValidator(
    file_format=FileFormat.ODF_1_3,
    schema_validation=True,
    semantic_validation=True,
    security_validation=False,
)

# security-core
security_core = OdfValidator(
    file_format=FileFormat.ODF_1_3,
    schema_validation=True,
    semantic_validation=True,
    security_validation=True,
    verify_cryptography=False,  # set True when crypto backend is available
)
```

### ODF Benchmarking

```bash
# Benchmark an ODF file (5 iterations by default)
python scripts/odf/benchmark_validation.py document.odt

# More iterations, with security checks
python scripts/odf/benchmark_validation.py document.odt --iterations 20 --security

# Foundation-only (skip schema/semantic)
python scripts/odf/benchmark_validation.py document.odt --no-schema --no-semantic
```

Reports avg/min/max/P95 with per-phase breakdown (package_structure, xml_parse, schema, semantic, security).

OOXML benchmark: `python scripts/benchmark_validation.py presentation.pptx`

### Known ODF Limitations

- Full OASIS conformance parity is not yet complete.
- Schema-core requires caller-provided Relax NG routes (`schema_routes`).
- Security-core validates structure/policy, not full cryptographic trust by default.
- CLI `--odf-level` only applies when the selected/auto-detected validator is ODF.

### ODF Reference Calibration

Compare Python results against external validators (ODF Toolkit, OPF) using the scripts in `scripts/odf/`:

| Script | Purpose |
|--------|---------|
| `run_reference_validators.py` | Run Python + external validators on pinned corpus |
| `compare_reference_results.py` | Diff results into mismatch families |
| `check_reference_drift.py` | Enforce drift policy against baseline |
| `bootstrap_reference_validators.py` | Auto-build external validator commands |

CI workflow: `.github/workflows/odf-reference-calibration.yml` — builds ODF Toolkit and OPF at runtime via Maven/Docker.

Set command templates via `--odf-toolkit-cmd` / `--opf-cmd` or env vars `ODF_TOOLKIT_CMD` / `OPF_ODF_VALIDATOR_CMD`. Placeholders: `{file}`, `{file_dir}`, `{file_name}`, `{file_stem}`, `{file_suffix}`.

## Open XML SDK (Standalone)

Run the .NET SDK validator separately (requires .NET SDK 8.x or Docker):

```bash
dotnet run --project scripts/sdk_check/sdk_check.csproj -- /path/to/file.pptx
dotnet run --project scripts/sdk_compare/OpenXmlSdkValidator.csproj -- /path/to/file.pptx  # JSON

# Via Docker
docker run --rm -v "$PWD:/work" -w /work mcr.microsoft.com/dotnet/sdk:8.0 \
  dotnet run --project scripts/sdk_check/sdk_check.csproj -- /work/path/to/file.pptx
```

Supports PPTX/DOCX/XLSX and variants. Configured for Office 2019.

## GitHub Action

Validate Office files in your PRs automatically:

```yaml
# .github/workflows/validate-office-files.yml
name: Validate Office Files
on: [pull_request]

jobs:
  validate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"
      - uses: BramAlkema/openxml-audit@main
        with:
          changed-only: "true"  # only validate files changed in the PR
```

Options:

| Input | Default | Description |
|-------|---------|-------------|
| `path` | `.` | Directory or file to validate |
| `format` | `Office2019` | Office version to validate against |
| `changed-only` | `false` | Only validate files changed in the PR |
| `recursive` | `true` | Search subdirectories |
| `max-errors` | `100` | Maximum errors per file |

## Pre-commit Hook

```yaml
# .pre-commit-config.yaml
repos:
  - repo: https://github.com/BramAlkema/openxml-audit
    rev: v0.4.2
    hooks:
      - id: openxml-audit
```

Validates any `.pptx`, `.docx`, `.xlsx`, `.odt`, `.ods`, or `.odp` file before commit.

## Examples

Ready-to-run scripts in [`examples/`](examples/):

| Script | Description |
|--------|-------------|
| [`validate_python_pptx.py`](examples/validate_python_pptx.py) | Generate a PPTX with python-pptx and validate it |
| [`validate_openpyxl.py`](examples/validate_openpyxl.py) | Generate an XLSX with openpyxl and validate it |
| [`validate_odf.py`](examples/validate_odf.py) | Validate an ODF file (ODT/ODS/ODP) |
| [`ci_validation.py`](examples/ci_validation.py) | Validate all Office files in a directory (CI-ready, OOXML + ODF) |

## CI Workflows

| Workflow | Trigger | Purpose |
|----------|---------|---------|
| `parity-gate.yml` | PR / push | Enforce OOXML parity + perf budget against SDK baseline |
| `calibrate-parity.yml` | Weekly / dispatch | Calibrate against Open XML SDK upstream |
| `sdk-update.yml` | Quarterly / dispatch | Track upstream SDK version changes |
| `odf-reference-calibration.yml` | Dispatch | Run ODF reference validators and drift checks |
| `validate-inputs.yml` | Push to `inputs/` | Validate dropped files with both Python and .NET SDK |
| `release.yml` | Tag push (`v*`) | Build and publish to PyPI |
| `pages.yml` | Push to `main` | Deploy documentation site |

OOXML parity details: `docs/parity_contract.md`. ODF reference contract: `docs/odf_validation_contract.md`.

## pytest Plugin

Fixtures are registered automatically — just `pip install openxml-audit` and use them:

```python
def test_my_presentation(assert_valid_pptx, tmp_path):
    output = tmp_path / "output.pptx"
    generate_pptx(output)
    assert_valid_pptx(output)  # fails with detailed errors if invalid

def test_my_document(assert_valid_docx, tmp_path):
    output = tmp_path / "output.docx"
    generate_docx(output)
    assert_valid_docx(output)

def test_my_spreadsheet(assert_valid_xlsx, tmp_path):
    output = tmp_path / "output.xlsx"
    generate_xlsx(output)
    assert_valid_xlsx(output)

def test_odf_file(assert_valid_odf, tmp_path):
    output = tmp_path / "output.odt"
    generate_odt(output)
    assert_valid_odf(output)
```

CLI options:

```bash
# Validate against a specific Office version
pytest --openxml-format Office2007

# Limit errors collected per file
pytest --openxml-max-errors 50
```

Available fixtures: `openxml_validator`, `assert_valid_pptx`, `assert_valid_docx`, `assert_valid_xlsx`, `assert_valid_odf`.

## Integration Helpers

```python
# Context manager
from openxml_audit import validation_context

with validation_context(raise_on_invalid=True) as validator:
    result = validator.validate("presentation.pptx")

# Decorator — validate after save
from openxml_audit import validate_on_save

@validate_on_save(raise_on_invalid=True)
def create_presentation(output_path: str) -> None:
    Presentation().save(output_path)

# Decorator — require valid input
from openxml_audit import require_valid_pptx

@require_valid_pptx()
def process(input_path: str) -> dict: ...
```

## API Reference

### `OpenXmlValidator` / `OdfValidator`

```python
OpenXmlValidator(file_format=FileFormat.OFFICE_2019, max_errors=1000,
                 schema_validation=True, semantic_validation=True)

OdfValidator(file_format=FileFormat.ODF_1_3, max_errors=1000,
             schema_validation=True, semantic_validation=True,
             security_validation=False, strict=True)
```

Both expose:
- `validate(path) -> ValidationResult`
- `validate_with_timings(path) -> (ValidationResult, dict[str, float])`
- `is_valid(path) -> bool`

### `ValidationResult`

| Property | Type | Description |
|----------|------|-------------|
| `is_valid` | `bool` | No ERROR-severity issues |
| `errors` | `list[ValidationError]` | All errors and warnings |
| `error_count` / `warning_count` | `int` | Counts by severity |
| `file_path` | `str` | Validated file path |
| `file_format` | `FileFormat` | Version validated against |

### `ValidationError`

| Property | Type | Description |
|----------|------|-------------|
| `error_type` | `ValidationErrorType` | `PACKAGE`, `BINARY`, `SCHEMA`, `SEMANTIC`, `RELATIONSHIP`, `MARKUP_COMPATIBILITY` |
| `severity` | `ValidationSeverity` | `ERROR`, `WARNING`, `INFO` |
| `description` | `str` | Human-readable message |
| `part_uri` | `str \| None` | Affected part URI |
| `path` | `str \| None` | XPath to affected element |

### Supported Formats

| OOXML | ODF |
|-------|-----|
| `OFFICE_2007` through `MICROSOFT_365` (default: `OFFICE_2019`) | `ODF_1_2`, `ODF_1_3` (default: `ODF_1_3`) |

### Convenience Functions

- `validate_pptx(path) -> ValidationResult`
- `is_valid_pptx(path) -> bool`

## Works Well With

These libraries create Office files — openxml-audit checks them:

| Library | Format | Link |
|---------|--------|------|
| [python-pptx](https://github.com/scanny/python-pptx) | PPTX | Create and update PowerPoint files |
| [python-docx](https://github.com/python-openxml/python-docx) | DOCX | Create and update Word files |
| [openpyxl](https://openpyxl.readthedocs.io/) | XLSX | Create and update Excel files |

```python
from pptx import Presentation
from openxml_audit import validate_pptx

Presentation().save("output.pptx")

result = validate_pptx("output.pptx")
if not result.is_valid:
    print(f"{result.error_count} issues found")
```

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for dev setup and guidelines.

## Looking for Maintainers

This project is actively looking for co-maintainers — especially people working with:

- Office file generation pipelines (python-pptx, python-docx, openpyxl)
- ODF tooling and OASIS conformance
- Open XML SDK internals

If you're interested, open an issue or reach out.

## Funding

If this project saves you time, consider sponsoring its development:

[![GitHub Sponsors](https://img.shields.io/badge/sponsor-GitHub%20Sponsors-ea4aaa)](https://github.com/sponsors/BramAlkema)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for a full list of changes by version.

## License

[MIT](LICENSE)

## Acknowledgments

Based on the validation logic from Microsoft's [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) for .NET.
