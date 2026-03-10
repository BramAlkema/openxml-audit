# OpenXML Audit

[![PyPI](https://img.shields.io/pypi/v/openxml-audit)](https://pypi.org/project/openxml-audit/)
[![Python](https://img.shields.io/pypi/pyversions/openxml-audit)](https://pypi.org/project/openxml-audit/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![CI](https://github.com/BramAlkema/openxml-audit/actions/workflows/parity-gate.yml/badge.svg)](https://github.com/BramAlkema/openxml-audit/actions/workflows/parity-gate.yml)

Validate OOXML (PPTX/DOCX/XLSX) and ODF files in pure Python — no .NET required.

A Python port of Microsoft's [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) validation logic. Check whether generated or modified Office files will open cleanly, directly from Python scripts, CI pipelines, or anywhere .NET isn't practical.

Also supports OASIS OpenDocument Format (ODT/ODS/ODP) with staged conformance levels.

## Features

- **OOXML Validation**: Package structure, schema, semantic, and format-specific checks for PPTX/DOCX/XLSX — matching the Open XML SDK's validation without the .NET dependency
- **ODF Validation**: Staged conformance levels — foundation, schema-core (Relax NG), semantic-core, and security-core for ODT/ODS/ODP
- **Multiple Output Formats**: Text, JSON, and XML output
- **Performance Tooling**: Per-phase timing breakdown, benchmark scripts for both OOXML and ODF
- **Flexible Integration**: Context managers, decorators, and pytest fixtures

## Installation

```bash
pip install openxml-audit
```

Or install from source:

```bash
git clone https://github.com/yourusername/openxml-audit.git
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

## CI Workflows

| Workflow | Trigger | Purpose |
|----------|---------|---------|
| `parity-gate.yml` | PR / push | Enforce OOXML parity + perf budget against SDK baseline |
| `calibrate-parity.yml` | Weekly / dispatch | Calibrate against Open XML SDK upstream |
| `odf-reference-calibration.yml` | Dispatch | Run ODF reference validators and drift checks |
| `validate-inputs.yml` | Push to `inputs/` | Validate dropped files with both Python and .NET SDK |

OOXML parity details: `docs/parity_contract.md`. ODF reference contract: `docs/odf_validation_contract.md`.

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

# pytest fixtures (add to conftest.py)
from openxml_audit.helpers import pytest_openxml_audit, pytest_assert_valid_pptx

openxml_audit = pytest_openxml_audit()
assert_valid_pptx = pytest_assert_valid_pptx()
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
| `error_type` | `ValidationErrorType` | `PACKAGE`, `SCHEMA`, `SEMANTIC`, `RELATIONSHIP` |
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

## License

[MIT](LICENSE)

## Acknowledgments

Based on the validation logic from Microsoft's [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) for .NET.
