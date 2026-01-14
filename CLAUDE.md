# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**openxml-audit** is a Python port of Microsoft's Open XML SDK validation. It validates OOXML (PPTX/DOCX/XLSX) and ODF files to determine if they will open successfully in their target apps.

## Commands

```bash
# Install dependencies
pip install -e ".[dev]"

# Run validator on a file
openxml-audit presentation.pptx

# Run with options
openxml-audit --format office2019 --output json presentation.pptx
openxml-audit --recursive ./presentations/

# Run tests
pytest

# Run linting
ruff check src/
ruff format src/

# Type checking
mypy src/openxml_audit
```

## Architecture

```
src/openxml_audit/
├── __init__.py         # Public API exports
├── validator.py        # OpenXmlValidator - main entry point
├── package.py          # OpenXmlPackage - OPC/ZIP handling
├── parts.py            # OpenXmlPart - document parts
├── relationships.py    # Relationship parsing and resolution
├── namespaces.py       # Open XML namespace constants
├── errors.py           # ValidationError, ValidationResult types
├── cli.py              # Click CLI interface
├── schema/             # Schema validation (particle validators, type validators)
└── semantic/           # Semantic validation (attribute, relationship, reference constraints)
```

### Key Classes

- **`OpenXmlValidator`**: Main validator class. Use `validate()` for detailed results or `is_valid()` for quick check.
- **`OpenXmlPackage`**: Handles OOXML as OPC packages (ZIP with XML). Parses `[Content_Types].xml` and `_rels/.rels`.
- **`OpenXmlPart`**: Represents a part in the package. Lazy-loads XML content.
- **`RelationshipCollection`**: Manages relationships between parts.

### Validation Phases

1. **Package Structure**: Valid ZIP, `[Content_Types].xml`, `_rels/.rels`, main document exists
2. **Presentation Structure**: `presentation.xml` parses, has slide masters
3. **Slide Validation**: All slides exist and parse, have layout relationships
4. **Relationship Integrity**: All internal relationships point to existing parts

## Open XML Concepts

- **OPC**: Open Packaging Conventions - PPTX is a ZIP containing XML parts
- **Part**: Individual XML file in the package (e.g., `/ppt/slides/slide1.xml`)
- **Relationship**: Links between parts, stored in `_rels/` directories
- **Content Type**: MIME type for each part, defined in `[Content_Types].xml`

## Adding New Validation

To add schema/semantic validation, implement in `schema/` or `semantic/` directories and integrate into `OpenXmlValidator._validate_*` methods.
