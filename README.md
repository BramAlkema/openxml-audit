# OpenXML Audit

A Python library for validating Open XML (OOXML) and ODF files against schema and semantic rules. Determine if Office files (PPTX/DOCX/XLSX) or ODF files will open cleanly in their target apps.

This is a Python port of the validation logic from Microsoft's [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

## Features

- **Package Validation**: Validates OPC (Open Packaging Conventions) structure
- **Schema Validation**: Validates XML content against Open XML schema constraints
- **Semantic Validation**: Validates relationships, references, and cross-document constraints
- **OOXML-Specific Validation**: Validates presentation, wordprocessing, and spreadsheet structure
- **Multiple Output Formats**: Text, JSON, and XML output
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

## Open XML SDK (Standalone)

Run the .NET Open XML SDK validator separately from the Python package. These helpers live under `scripts/` and only require the .NET SDK (or Docker).

### Local .NET SDK

1. Install the .NET SDK (8.x recommended).
2. From the repo root, run:

```bash
# Plain text output
dotnet run --project scripts/sdk_check/sdk_check.csproj -- /path/to/file.pptx

# JSON output (useful for diffs)
dotnet run --project scripts/sdk_compare/OpenXmlSdkValidator.csproj -- /path/to/file.pptx
```

Notes:
- The SDK validator is configured for Office 2019 (see `scripts/sdk_check/Program.cs`).
- It supports PPTX/POTX/PPSX, DOCX/DOTX, and XLSX/XLTX inputs.
- Mono is not supported; use the .NET SDK or Docker.

### Docker (no local .NET install)

```bash
docker run --rm -v "$PWD:/work" -w /work mcr.microsoft.com/dotnet/sdk:8.0 \
  dotnet run --project scripts/sdk_check/sdk_check.csproj -- /work/path/to/file.pptx
```

For JSON output, swap the project to `scripts/sdk_compare/OpenXmlSdkValidator.csproj`.

## GitHub Actions Validation

Drop files into `inputs/` and push to GitHub. The workflow will:
- run the Python validator and the Open XML SDK validator
- upload JSON reports as artifacts
- post a summary in the GitHub Actions job summary

Workflow file: `.github/workflows/validate-inputs.yml`.

## Integration Helpers

### Context Manager

```python
from openxml_audit import validation_context

# Basic usage
with validation_context() as validator:
    result = validator.validate("presentation.pptx")
    print(f"Valid: {result.is_valid}")

# Raise exception on invalid files
with validation_context(raise_on_invalid=True) as validator:
    result = validator.validate("presentation.pptx")  # Raises ValueError if invalid
```

### Decorator for python-pptx

Validate PPTX files created by python-pptx after saving:

```python
from pptx import Presentation
from openxml_audit import validate_on_save

@validate_on_save(raise_on_invalid=True)
def create_presentation(output_path: str) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    title.text = "Hello World"
    prs.save(output_path)

# Will validate after save, raises ValueError if invalid
create_presentation("output.pptx")
```

Validate input files before processing:

```python
from openxml_audit import require_valid_pptx

@require_valid_pptx()
def process_presentation(input_path: str) -> dict:
    # This only runs if input is valid
    prs = Presentation(input_path)
    return {"slides": len(prs.slides)}

# Raises ValueError if input is invalid
result = process_presentation("input.pptx")
```

### pytest Fixtures

Add to your `conftest.py`:

```python
from openxml_audit.helpers import (
    pytest_openxml_audit,
    pytest_valid_pptx_path,
    pytest_assert_valid_pptx,
)

# Register fixtures
openxml_audit = pytest_openxml_audit()
valid_pptx_path = pytest_valid_pptx_path()
assert_valid_pptx = pytest_assert_valid_pptx()
```

Use in tests:

```python
def test_my_generator(assert_valid_pptx, tmp_path):
    output = tmp_path / "output.pptx"
    generate_my_pptx(output)
    assert_valid_pptx(output)  # Fails with detailed error message if invalid

def test_with_validator(openxml_audit, tmp_path):
    output = tmp_path / "output.pptx"
    create_pptx(output)
    result = openxml_audit.validate(output)
    assert result.is_valid
```

## Validation Phases

The validator runs through multiple validation phases:

1. **Package Structure**: Validates ZIP structure, content types, and package relationships
2. **Presentation Structure**: Validates presentation.xml and slide/master references
3. **Slide Validation**: Validates slide XML structure and layout relationships
4. **Relationship Integrity**: Validates all relationships point to existing parts
5. **Schema Validation**: Validates XML content against element constraints
6. **Semantic Validation**: Validates cross-document references and constraints
7. **OOXML-Specific Validation**: Validates themes, masters, layouts, and slides (plus Word/Excel structure)

## Office Format Versions

Validate against different Office versions:

```python
from openxml_audit import OpenXmlValidator, FileFormat

# Available versions
FileFormat.OFFICE_2007    # Office 2007 (ECMA-376 1st Edition)
FileFormat.OFFICE_2010    # Office 2010
FileFormat.OFFICE_2013    # Office 2013
FileFormat.OFFICE_2016    # Office 2016
FileFormat.OFFICE_2019    # Office 2019 (default)
FileFormat.MICROSOFT_365  # Microsoft 365

validator = OpenXmlValidator(file_format=FileFormat.OFFICE_2007)
```

## Error Types

Errors are categorized by type:

- `PACKAGE`: OPC package structure issues (missing parts, invalid ZIP)
- `SCHEMA`: XML schema violations (missing elements, invalid values)
- `SEMANTIC`: Semantic constraint violations (broken references, invalid IDs)
- `RELATIONSHIP`: Relationship issues (missing targets, wrong types)

And by severity:

- `ERROR`: Critical issues that will prevent the file from opening
- `WARNING`: Issues that may cause problems but won't prevent opening
- `INFO`: Informational messages about potential issues

## API Reference

### Classes

#### `OpenXmlValidator`

Main validator class.

```python
OpenXmlValidator(
    file_format: FileFormat = FileFormat.OFFICE_2019,
    max_errors: int = 1000,
    schema_validation: bool = True,
    semantic_validation: bool = True,
)
```

**Methods:**
- `validate(path: str | Path) -> ValidationResult`: Validate an OOXML file
- `is_valid(path: str | Path) -> bool`: Quick check if file is valid

#### `ValidationResult`

Result of validation.

**Properties:**
- `is_valid: bool`: True if no errors found
- `errors: list[ValidationError]`: List of all errors and warnings
- `error_count: int`: Number of ERROR severity issues
- `warning_count: int`: Number of WARNING severity issues
- `file_path: str`: Path to validated file
- `file_format: FileFormat`: Office version validated against

#### `ValidationError`

Individual validation error.

**Properties:**
- `error_type: ValidationErrorType`: Category of error
- `severity: ValidationSeverity`: ERROR, WARNING, or INFO
- `description: str`: Human-readable error description
- `part_uri: str | None`: URI of the affected part
- `path: str | None`: XPath to the affected element
- `node: str | None`: Name of the affected node/attribute

### Functions

- `validate_pptx(path: str | Path) -> ValidationResult`: Validate with default options
- `is_valid_pptx(path: str | Path) -> bool`: Quick validity check

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

This project is based on the validation logic from Microsoft's [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) for .NET.
