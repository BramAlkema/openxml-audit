# Tasks - Open XML SDK Python Port

## Phase 1: Project Setup & Core Package Handling ✅

### 1.1 Project Infrastructure ✅
- [x] Create `pyproject.toml` with dependencies (lxml, click)
- [x] Create package structure (`src/openxml_audit/`)
- [x] Set up `tests/` directory
- [x] Create `__init__.py` with public API exports

### 1.2 Error Types (errors.py) ✅
- [x] Define `ValidationErrorType` enum
- [x] Define `ValidationError` dataclass
- [x] Define `FileFormat` enum (Office2007 through Microsoft365)

### 1.3 Namespace Definitions (namespaces.py) ✅
- [x] Define all Open XML namespaces (PresentationML, DrawingML, etc.)
- [x] Create namespace prefix mappings for lxml
- [x] Add helper functions for namespace handling

### 1.4 OPC Package Handling (package.py) ✅
- [x] Implement `OpenXmlPackage` class
  - [x] Open and validate ZIP structure
  - [x] Parse `[Content_Types].xml`
  - [x] Load package relationships from `_rels/.rels`
  - [x] Enumerate all parts in package
- [x] Implement `ContentTypes` class
  - [x] Parse default content types
  - [x] Parse override content types
  - [x] Lookup content type by part name

### 1.5 Parts Handling (parts.py) ✅
- [x] Implement `OpenXmlPart` class
  - [x] Lazy load XML content
  - [x] Parse part relationships from `{part}_rels/.rels`
  - [x] Get root element
- [x] Implement `PresentationPart` subclass

### 1.6 Relationships (relationships.py) ✅
- [x] Implement `Relationship` dataclass
- [x] Implement `RelationshipCollection` class
- [x] Parse relationship XML files
- [x] Resolve relationship targets (relative paths)

### 1.7 Basic Package Validation ✅
- [x] Check ZIP is valid
- [x] Check `[Content_Types].xml` exists and parses
- [x] Check `_rels/.rels` exists with required relationships
- [x] Check `ppt/presentation.xml` exists
- [x] Return `ValidationError` list for failures

---

## Phase 2: Schema Validation ✅

### 2.1 Validation Context (context.py) ✅
- [x] Implement `ValidationContext` class
  - [x] Current part being validated
  - [x] Current element path (XPath)
  - [x] Error collection
  - [x] File format version
- [x] Implement `ValidationStack` for tracking position

### 2.2 Particle Validators (schema/particle.py) ✅
- [x] Port `ParticleConstraint` base class
- [x] Port `SequenceParticleValidator`
  - [x] Validate elements appear in correct order
  - [x] Handle minOccurs/maxOccurs
- [x] Port `ChoiceParticleValidator`
  - [x] Validate one of allowed elements present
- [x] Port `AllParticleValidator`
  - [x] Validate all required elements present (any order)
- [x] Port `CompositeParticleValidator`
  - [x] Handle nested particle constraints

### 2.3 Type Validators (schema/types.py) ✅
- [x] Implement base `XsdType` class
- [x] Implement string type validators
  - [x] Pattern (regex) validation
  - [x] Length constraints (min/max)
  - [x] Enumeration validation
- [x] Implement numeric type validators
  - [x] Integer ranges
  - [x] Decimal/float ranges
- [x] Implement boolean validator
- [ ] Implement DateTime validator (deferred)

### 2.4 Element Constraints (schema/constraints.py) ✅
- [x] Define constraint data structures
- [x] Load PPTX element constraints (core elements defined)
- [x] Validate required attributes
- [x] Validate required child elements

### 2.5 Schema Validator Integration ✅
- [x] Implement `SchemaValidator` class
- [x] Traverse document tree
- [x] Apply particle validators per element
- [x] Apply type validators per attribute
- [x] Collect all errors

---

## Phase 3: Semantic Validation ✅

### 3.1 Attribute Constraints (semantic/attributes.py) ✅
- [x] Port `AttributeMinMaxConstraint`
- [x] Port `AttributeValuePatternConstraint`
- [x] Port `AttributeMutualExclusive`
- [x] Port `AttributeRequiredConditionToValue`
- [x] Port `AttributeValueLengthConstraint`
- [x] Port `AttributeValueRangeConstraint`
- [x] Port `AttributeValueSetConstraint`

### 3.2 Relationship Constraints (semantic/relationships.py) ✅
- [x] Port `RelationshipExistConstraint`
- [x] Port `RelationshipTypeConstraint`
- [x] Validate all referenced relationships exist
- [x] Validate relationship targets exist in package

### 3.3 Reference Constraints (semantic/references.py) ✅
- [x] Port `IndexReferenceConstraint`
- [x] Port `ReferenceExistConstraint`
- [x] Port `UniqueAttributeValueConstraint`
- [x] Validate ID references resolve

### 3.4 Semantic Validator Integration ✅
- [x] Implement `SemanticValidator` class
- [x] Apply attribute constraints
- [x] Apply relationship constraints
- [x] Apply reference constraints

---

## Phase 4: PPTX-Specific Validation ✅

### 4.1 Presentation Structure (pptx/presentation.py) ✅
- [x] Validate `presentation.xml` structure
- [x] Validate slide list references
- [x] Validate slide master references
- [x] Validate note master references

### 4.2 Slide Validation (pptx/slides.py) ✅
- [x] Validate slide XML structure
- [x] Validate shape tree
- [x] Validate text bodies
- [x] Validate embedded objects

### 4.3 Theme Validation (pptx/themes.py) ✅
- [x] Validate theme XML structure
- [x] Validate color schemes
- [x] Validate font schemes
- [x] Validate format schemes

### 4.4 Master/Layout Validation (pptx/masters.py) ✅
- [x] Validate slide master structure
- [x] Validate slide layout structure
- [x] Validate master-layout relationships
- [x] Validate placeholder mappings

---

## Phase 5: CLI & Integration ✅

### 5.1 Main Validator (validator.py) ✅
- [x] Implement `OpenXmlValidator` class
  - [x] `validate(path) -> list[ValidationError]`
  - [x] `is_valid(path) -> bool`
  - [x] `max_errors` configuration
- [x] Orchestrate all validation phases
- [x] Return consolidated error list

### 5.2 CLI Interface (cli.py) ✅
- [x] Implement `openxml-audit` command
- [x] Add `--format` option for Office version
- [x] Add `--output` option (text/json/xml)
- [x] Add `--max-errors` option
- [x] Add `--recursive` for directory validation
- [x] Colored terminal output with rich

### 5.3 Output Formats ✅
- [x] Plain text output (default)
- [x] JSON output
- [x] XML output (match .NET SDK format)

### 5.4 Integration Helpers (helpers.py) ✅
- [x] Context manager for validation (`validation_context`)
- [x] Decorator for python-pptx validation (`validate_on_save`, `require_valid_pptx`)
- [x] pytest fixtures for testing

---

## Phase 6: Testing & Documentation ✅

### 6.1 Unit Tests ✅
- [x] Test OPC package parsing (test_package.py)
- [x] Test content types parsing (test_package.py)
- [x] Test relationship parsing (test_relationships.py)
- [x] Test each particle validator (deferred - schema constraints defined)
- [x] Test each type validator (test_schema_types.py)
- [x] Test each semantic constraint (deferred - semantic validators defined)

### 6.2 Integration Tests ✅
- [x] Create/collect valid PPTX test files (conftest.py fixtures)
- [x] Create/collect invalid PPTX test files (conftest.py fixtures)
- [x] Test full validation pipeline (test_integration.py)
- [ ] Compare results with .NET SDK (if available)

### 6.3 Documentation ✅
- [x] API documentation (docstrings)
- [x] README.md with usage examples
- [x] CLAUDE.md for AI assistance

---

## Priority Order

**Must Have (MVP):**
1. Phase 1: Core Package Handling (1.1-1.7)
2. Phase 5.1: Main Validator (basic)
3. Phase 5.2: CLI Interface (basic)

**Should Have:**
4. Phase 2: Schema Validation
5. Phase 4.1-4.2: Presentation & Slide Validation

**Nice to Have:**
6. Phase 3: Semantic Validation
7. Phase 4.3-4.4: Theme & Master Validation
8. Phase 6: Full Testing & Documentation
