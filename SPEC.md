# Open XML SDK Python Port - Specification

## Overview

Port Microsoft's Open XML SDK validation framework to Python to validate PPTX files and determine if they will open in PowerPoint.

## Problem Statement

- `python-pptx` and similar libraries can produce PPTX files that pass basic XML validation but fail to open in PowerPoint
- Microsoft's Open XML SDK (.NET) has comprehensive validation but requires .NET runtime
- No Python-native solution exists for thorough PPTX validation

## Goals

1. **Primary**: Validate if a PPTX file will open in PowerPoint
2. **Secondary**: Provide detailed error information when validation fails
3. **Tertiary**: Maintain compatibility with the .NET SDK's validation behavior

## Non-Goals

- Document creation/manipulation (use python-pptx for that)
- Full port of Open XML SDK (only validation components)
- Support for DOCX/XLSX initially (PPTX only, can extend later)

---

## Technical Specification

### Architecture

```
openxml_audit/
├── __init__.py
├── validator.py          # Main OpenXmlValidator class
├── package.py            # OpenXmlPackage - ZIP/OPC handling
├── parts.py              # OpenXmlPart - document parts
├── relationships.py      # Relationship handling
├── context.py            # ValidationContext
├── errors.py             # ValidationError types
├── schema/
│   ├── __init__.py
│   ├── particle.py       # Particle validators (sequence, choice, all)
│   ├── types.py          # XSD type validation
│   └── constraints.py    # Element constraints
├── semantic/
│   ├── __init__.py
│   ├── attributes.py     # Attribute constraints
│   ├── relationships.py  # Relationship constraints
│   └── references.py     # Reference constraints
└── namespaces.py         # Open XML namespace definitions
```

### Core Components

#### 1. OpenXmlValidator (validator.py)

Main entry point for validation.

```python
class OpenXmlValidator:
    def __init__(self, file_format: FileFormat = FileFormat.Office2019):
        ...

    def validate(self, path: str | Path) -> list[ValidationError]:
        """Validate a PPTX file and return all errors."""
        ...

    def is_valid(self, path: str | Path) -> bool:
        """Quick check if file is valid."""
        ...
```

#### 2. OpenXmlPackage (package.py)

Handles OPC (Open Packaging Conventions) - PPTX is a ZIP with XML.

```python
class OpenXmlPackage:
    def __init__(self, path: str | Path):
        ...

    @property
    def content_types(self) -> ContentTypes:
        """Parse [Content_Types].xml"""
        ...

    @property
    def parts(self) -> dict[str, OpenXmlPart]:
        """All document parts"""
        ...

    @property
    def relationships(self) -> list[Relationship]:
        """Package-level relationships from _rels/.rels"""
        ...
```

#### 3. Schema Validation (schema/)

Validates XML structure against Open XML schema rules.

**Particle Validators** - Validate element sequences:
- `SequenceParticleValidator` - Elements must appear in order
- `ChoiceParticleValidator` - One of several options
- `AllParticleValidator` - All elements, any order
- `CompositeParticleValidator` - Nested particles

**Type Validators** - Validate attribute/element values:
- String patterns (regex)
- Numeric ranges
- Enumerations
- Boolean values

#### 4. Semantic Validation (semantic/)

Validates logical constraints beyond schema.

**Attribute Constraints:**
- `AttributeMinMaxConstraint` - Value range limits
- `AttributeValuePatternConstraint` - Regex patterns
- `AttributeMutualExclusive` - Can't have both attributes
- `AttributeRequiredConditionToValue` - Required if another attr has value

**Relationship Constraints:**
- `RelationshipExistConstraint` - Referenced relationships must exist
- `RelationshipTypeConstraint` - Relationship type must match

**Reference Constraints:**
- `IndexReferenceConstraint` - Index references must be valid
- `ReferenceExistConstraint` - Referenced elements must exist

### Dependencies

```toml
[project]
dependencies = [
    "lxml>=5.0.0",      # XML parsing and XPath
    "click>=8.1.0",     # CLI interface
]
```

### File Format Versions

Support validation against different Office versions:

```python
class FileFormat(Enum):
    Office2007 = "office2007"  # ECMA-376 1st edition
    Office2010 = "office2010"  # ECMA-376 2nd edition
    Office2013 = "office2013"
    Office2016 = "office2016"
    Office2019 = "office2019"
    Office2021 = "office2021"
    Microsoft365 = "microsoft365"
```

---

## Implementation Phases

### Phase 1: Core Package Handling
- [ ] ZIP extraction and OPC parsing
- [ ] Content types parsing (`[Content_Types].xml`)
- [ ] Relationship parsing (`_rels/*.rels`)
- [ ] Basic part enumeration
- [ ] Namespace definitions

**Validation checks:**
- Is it a valid ZIP?
- Does `[Content_Types].xml` exist and parse?
- Does `_rels/.rels` exist with required relationships?
- Does `ppt/presentation.xml` exist?

### Phase 2: Schema Validation
- [ ] Port particle validators from .NET SDK
- [ ] Port XSD type validators
- [ ] Implement element constraint checking
- [ ] Validate required elements exist

**Source files to port:**
- `ParticleValidator.cs` → `schema/particle.py`
- `SequenceParticleValidator.cs`
- `ChoiceParticleValidator.cs`
- `SchemaTypeValidator.cs` → `schema/types.py`

### Phase 3: Semantic Validation
- [ ] Port attribute constraints
- [ ] Port relationship constraints
- [ ] Port reference constraints

**Source files to port:**
- `AttributeMinMaxConstraint.cs` → `semantic/attributes.py`
- `RelationshipExistConstraint.cs` → `semantic/relationships.py`
- `ReferenceExistConstraint.cs` → `semantic/references.py`

### Phase 4: PPTX-Specific Validation
- [ ] Presentation structure validation
- [ ] Slide validation
- [ ] Theme validation
- [ ] Master/layout relationships

### Phase 5: CLI & Integration
- [ ] CLI tool (`openxml-audit` command)
- [ ] JSON/XML output formats
- [ ] Integration with python-pptx workflow

---

## Validation Error Format

```python
@dataclass
class ValidationError:
    error_type: ValidationErrorType
    description: str
    path: str          # XPath to element
    part: str          # Part URI (e.g., "/ppt/slides/slide1.xml")
    node: str | None   # Element/attribute name
    related_node: str | None

class ValidationErrorType(Enum):
    SCHEMA = "schema"           # Schema violation
    SEMANTIC = "semantic"       # Semantic constraint violation
    PACKAGE = "package"         # OPC package error
    RELATIONSHIP = "relationship"
    MARKUP_COMPATIBILITY = "markup_compatibility"
```

---

## Testing Strategy

1. **Unit tests** - Individual validators
2. **Integration tests** - Full validation of known-good/bad files
3. **Comparison tests** - Compare results with .NET SDK output

### Test Files Needed
- Valid PPTX files from different Office versions
- Invalid PPTX files with known errors
- Edge cases (empty slides, missing parts, etc.)

---

## CLI Interface

```bash
# Basic validation
openxml-audit presentation.pptx

# Specify Office version
openxml-audit --format office2019 presentation.pptx

# Output as JSON
openxml-audit --output json presentation.pptx

# Validate directory
openxml-audit --recursive ./presentations/

# Limit errors
openxml-audit --max-errors 10 presentation.pptx
```

---

## Open Questions

1. **Schema data source**: Should we extract constraint data from .NET SDK or define independently?
2. **Version differences**: How to handle version-specific validation rules?
3. **Performance**: Lazy vs eager loading of parts?

---

## References

- [.NET Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- [ECMA-376 Standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [OPC Specification](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview)
- [OpenXmlValidator API](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator)
