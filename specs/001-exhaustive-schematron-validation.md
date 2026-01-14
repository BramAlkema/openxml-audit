# Spec: Exhaustive Schematron Validation (100% Coverage)

## Overview

Achieve 100% coverage of the 948 schematron rules from Microsoft's Open XML SDK by:
1. Wiring up the 742 already-interpretable rules
2. Extending pattern matching to cover the remaining 206 rules

## Current State

```
Schema Validation:     100% (3,701 elements from SDK JSON)
Semantic Validation:     0% (6 hardcoded rules, 948 SDK rules unused)
```

## Target State

```
Schema Validation:     100%
Semantic Validation:   100% (948 rules active)
```

---

## Phase 1: Wire Interpretable Rules (742 rules, 78%)

### 1.1 Schematron-to-Constraint Bridge

**File:** `src/openxml_audit/codegen/schematron_bridge.py`

Create a bridge that converts parsed schematrons into `SemanticConstraint` instances:

```python
def create_constraint_from_schematron(rule: ParsedSchematron) -> SemanticConstraint | None:
    """Convert a parsed schematron to a semantic constraint."""
    match rule.rule_type:
        case SchematronType.ATTRIBUTE_VALUE_RANGE:
            return AttributeMinMaxConstraint(
                attribute=rule.attribute,
                min_value=rule.min_value,
                max_value=rule.max_value,
            )
        case SchematronType.ATTRIBUTE_VALUE_LENGTH:
            return AttributeStringLengthConstraint(
                attribute=rule.attribute,
                min_length=rule.min_length,
                max_length=rule.max_length,
            )
        # ... etc
```

### 1.2 Auto-Registration in SemanticValidator

**File:** `src/openxml_audit/semantic/validator.py`

Modify `create_pptx_semantic_validator()` to auto-register SDK rules:

```python
def create_pptx_semantic_validator() -> SemanticValidator:
    validator = SemanticValidator()

    # Load all interpretable schematron rules
    from openxml_audit.codegen.schematron_bridge import load_sdk_constraints
    for element_tag, constraint in load_sdk_constraints():
        validator.register_constraint(element_tag, constraint)

    return validator
```

### 1.3 New Constraint Classes Needed

| Rule Type | Count | Constraint Class | Status |
|-----------|-------|------------------|--------|
| ATTRIBUTE_VALUE_RANGE | 239 | `AttributeMinMaxConstraint` | Exists |
| UNIQUE_ATTRIBUTE | 213 | `UniqueAttributeConstraint` | **New** |
| ATTRIBUTE_VALUE_LENGTH | 186 | `AttributeStringLengthConstraint` | **New** |
| RELATIONSHIP_TYPE | 64 | `RelationshipTypeConstraint` | **New** |
| ELEMENT_REFERENCE | 23 | `ElementReferenceConstraint` | **New** |
| ATTRIBUTE_VALUE_PATTERN | 17 | `AttributePatternConstraint` | **New** |

### Tasks for Phase 1

- [ ] Create `UniqueAttributeConstraint` class
- [ ] Create `AttributeStringLengthConstraint` class
- [ ] Create `RelationshipTypeConstraint` class
- [ ] Create `ElementReferenceConstraint` class
- [ ] Create `AttributePatternConstraint` class
- [ ] Create `schematron_bridge.py` with conversion logic
- [ ] Update `SemanticValidator` to auto-load SDK rules
- [ ] Add tests for each new constraint type
- [ ] Verify 742 rules are active

---

## Phase 2: Extend Pattern Matching (206 rules, 22%)

### 2.1 Unknown Rule Categories

| Category | Count | Complexity | Approach |
|----------|-------|------------|----------|
| OR conditions | 67 | Medium | Parse OR logic |
| Other patterns | 56 | Variable | Individual analysis |
| Cross-part document() | 53 | High | Package-level validation |
| Inequality checks | 21 | Low | Simple != comparison |
| Multi-attribute AND | 9 | Medium | Compound constraints |

### 2.2 OR Conditions (67 rules)

**Example:**
```xpath
(@x:operator and @x:type = cells) or @x:type != cells
```

**Solution:** Create `OrConstraint` that wraps multiple sub-constraints:

```python
@dataclass
class OrConstraint(SemanticConstraint):
    """One of the sub-constraints must pass."""
    constraints: list[SemanticConstraint]

    def validate(self, element, context) -> bool:
        return any(c.validate(element, context) for c in self.constraints)
```

**Parser enhancement:**
```python
def parse_or_condition(test: str) -> OrConstraint:
    parts = re.split(r'\s+or\s+', test, flags=re.IGNORECASE)
    return OrConstraint([parse_condition(p) for p in parts])
```

### 2.3 Inequality Checks (21 rules)

**Example:**
```xpath
@x:guid != 00000000-0000-0000-0000-000000000000
```

**Solution:** Add `AttributeNotEqualConstraint`:

```python
@dataclass
class AttributeNotEqualConstraint(SemanticConstraint):
    attribute: str
    forbidden_value: str
```

### 2.4 Multi-Attribute AND (9 rules)

**Example:**
```xpath
@x:maxValue != NaN and @x:maxValue != INF and @x:maxValue != -INF
```

**Solution:** Create `AndConstraint`:

```python
@dataclass
class AndConstraint(SemanticConstraint):
    """All sub-constraints must pass."""
    constraints: list[SemanticConstraint]
```

### 2.5 Cross-Part Document Queries (53 rules)

**Example:**
```xpath
@x:cm < count(document('Part:/WorkbookPart/CellMetadataPart')//x:cellMetadata/x:bk) + 1
```

**Solution:** Package-level validation context:

```python
@dataclass
class CrossPartCountConstraint(SemanticConstraint):
    """Validate attribute against count from another part."""
    attribute: str
    target_part: str  # e.g., "WorkbookPart/CellMetadataPart"
    target_xpath: str  # e.g., "//x:cellMetadata/x:bk"
    comparison: str  # "<", "<=", "=", etc.

    def validate(self, element, context) -> bool:
        # Get target part from package
        target = context.package.get_part(self.target_part)
        if target is None:
            return True  # Part doesn't exist, skip

        # Count matching elements
        count = len(target.xml.xpath(self.target_xpath, namespaces=NS_MAP))

        # Compare
        attr_value = int(element.get(self.attribute, 0))
        return self._compare(attr_value, count)
```

### 2.6 Other Patterns (56 rules)

Analyze and categorize individually. Common patterns found:

| Pattern | Count | Solution |
|---------|-------|----------|
| `string-length(@x) >= 1` | ~10 | Already covered by LENGTH type |
| `@x = false` / `@x = true` | ~8 | `AttributeEqualsConstraint` |
| `@x and @y` (presence check) | ~12 | `AttributesPresentConstraint` |
| `@x < @y` (attribute comparison) | ~6 | `AttributeComparisonConstraint` |
| `@x:spt = N` (specific value) | ~8 | `AttributeEqualsConstraint` |
| Complex expressions | ~12 | Manual implementation |

### Tasks for Phase 2

- [ ] Create `OrConstraint` class
- [ ] Create `AndConstraint` class
- [ ] Create `AttributeNotEqualConstraint` class
- [ ] Create `AttributeEqualsConstraint` class
- [ ] Create `AttributesPresentConstraint` class
- [ ] Create `AttributeComparisonConstraint` class
- [ ] Create `CrossPartCountConstraint` class
- [ ] Enhance schematron parser for OR/AND patterns
- [ ] Enhance schematron parser for inequality
- [ ] Add package context to validation
- [ ] Manually implement remaining complex rules
- [ ] Add tests for Phase 2 constraints
- [ ] Verify all 948 rules are active

---

## Phase 3: Validation & Polish

### 3.1 Coverage Verification

Create a test that verifies all 948 rules are active:

```python
def test_all_schematrons_covered():
    from openxml_audit.codegen import get_schematron_registry
    reg = get_schematron_registry()

    # All rules should be convertible
    unconverted = []
    for rule in reg._rules:
        constraint = create_constraint_from_schematron(rule)
        if constraint is None:
            unconverted.append(rule)

    assert len(unconverted) == 0, f"{len(unconverted)} rules not converted"
```

### 3.2 Performance Optimization

- Cache constraint instances (already done via `@lru_cache`)
- Lazy-load schematrons only when semantic validation enabled
- Profile validation on large PPTX files

### 3.3 Error Message Quality

Ensure error messages from SDK rules are clear:

```python
# Bad: "Validation failed"
# Good: "Attribute 'x:windowWidth' value 999999999 exceeds maximum 2147483647"
```

---

## File Changes Summary

### New Files
- `src/openxml_audit/semantic/constraints/length.py`
- `src/openxml_audit/semantic/constraints/unique.py`
- `src/openxml_audit/semantic/constraints/pattern.py`
- `src/openxml_audit/semantic/constraints/relationship_type.py`
- `src/openxml_audit/semantic/constraints/compound.py` (Or/And)
- `src/openxml_audit/semantic/constraints/cross_part.py`
- `src/openxml_audit/codegen/schematron_bridge.py`
- `tests/test_schematron_constraints.py`

### Modified Files
- `src/openxml_audit/semantic/validator.py` - Auto-load SDK rules
- `src/openxml_audit/codegen/schematron_loader.py` - Enhanced parsing
- `src/openxml_audit/context.py` - Add package reference

---

## Success Criteria

1. All 948 schematron rules converted to constraints
2. All constraints registered in SemanticValidator
3. Test coverage for each constraint type
4. No performance regression (< 2x slower)
5. Clear error messages for all validation failures

---

## Estimated Effort

| Phase | Tasks | Complexity |
|-------|-------|------------|
| Phase 1 | 9 | Medium |
| Phase 2 | 12 | High |
| Phase 3 | 3 | Low |

**Total: 24 tasks**
