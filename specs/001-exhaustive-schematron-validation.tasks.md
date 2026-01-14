# Tasks: Exhaustive Schematron Validation (100% Coverage)

**Spec:** [001-exhaustive-schematron-validation.md](./001-exhaustive-schematron-validation.md)

---

## Phase 1: Wire Interpretable Rules (742 rules → 78%)

### 1.1 New Constraint Classes

- [x] **P1-1** Create `AttributeStringLengthConstraint` ✅ Already existed
  - File: `src/openxml_audit/semantic/constraints/length.py`
  - Validates: `string-length(@attr) >= min and <= max`
  - Test: `tests/test_semantic_constraints.py::TestStringLengthConstraint`
  - Covers: 186 rules

- [x] **P1-2** Create `UniqueAttributeConstraint` ✅ Already existed as `UniqueAttributeValueConstraint`
  - File: `src/openxml_audit/semantic/references.py`
  - Covers: 211/213 rules (99%)

- [x] **P1-3** Create `RelationshipTypeConstraint` ✅ Already existed
  - File: `src/openxml_audit/semantic/relationships.py`
  - Covers: 27/64 rules (42%)

- [x] **P1-4** Create `ElementReferenceConstraint` ✅ Already existed as `ReferenceExistConstraint`
  - File: `src/openxml_audit/semantic/references.py`
  - Covers: 0/23 rules (needs work in Phase 2)

- [x] **P1-5** Create `AttributePatternConstraint` ✅ Already existed as `AttributeValuePatternConstraint`
  - File: `src/openxml_audit/semantic/attributes.py`
  - Covers: 14/17 rules (82%)

### 1.2 Bridge & Integration

- [x] **P1-6** Create schematron-to-constraint bridge ✅
  - File: `src/openxml_audit/codegen/schematron_bridge.py`
  - Converts `ParsedSchematron` → `SemanticConstraint`
  - Handles XPath pattern conversion

- [x] **P1-7** Auto-register SDK constraints in SemanticValidator ✅
  - File: `src/openxml_audit/semantic/validator.py`
  - 676 constraints now auto-registered
  - 329 unique element tags

- [x] **P1-8** Add namespace resolution for schematron contexts ✅
  - Built into `schematron_bridge.py`
  - Uses namespace map from schema loader

### 1.3 Verification

- [x] **P1-9** Test Phase 1 coverage ✅
  - 673/948 rules converted (70%)
  - All 96 tests pass
  - Ready for Phase 2

---

## Phase 2: Extend Pattern Matching (206 rules → 22%)

### 2.1 Compound Constraints

- [x] **P2-1** Create `OrConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/compound.py`
  - Validates: Any sub-constraint passes
  - Covers: 22/26 rules (85%)

- [x] **P2-2** Create `AndConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/compound.py`
  - Validates: All sub-constraints pass
  - Covers: 1/1 rules (100%)

- [x] **P2-NEW** Create `ConditionalConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/compound.py`
  - Validates: If trigger attribute present, condition must hold
  - Covers: 16/17 rules (94%)

### 2.2 Simple Value Constraints

- [x] **P2-3** Create `AttributeNotEqualConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/equality.py`
  - Validates: `@attr != forbidden_value`
  - Covers: 21/21 rules (100%)

- [x] **P2-4** Create `AttributeEqualsConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/equality.py`
  - Validates: `@attr = expected_value`
  - Covers: 59/61 rules (97%)

- [x] **P2-5** Create `AttributesPresentConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/equality.py`
  - Validates: `@a and @b` (both attributes exist)
  - Covers: 14/14 rules (100%)

- [x] **P2-6** Create `AttributeComparisonConstraint` ✅
  - File: `src/openxml_audit/semantic/constraints/equality.py`
  - Validates: `@a < @b`, `@a <= @b`, etc.
  - Covers: 6/6 rules (100%)

### 2.3 Cross-Part Validation

- [ ] **P2-7** Add package reference to ValidationContext
  - File: `src/openxml_audit/context.py`
  - Add: `package: OpenXmlPackage` field
  - Allow constraints to access other parts

- [ ] **P2-8** Create `CrossPartCountConstraint`
  - File: `src/openxml_audit/semantic/constraints/cross_part.py`
  - Validates: `@attr < count(document('Part:X')//xpath)`
  - Covers: 53 rules (currently skipped)

### 2.4 Parser Enhancements

- [x] **P2-9** Enhance schematron parser for OR conditions ✅
  - File: `src/openxml_audit/codegen/schematron_loader.py`
  - Parse: `(cond1) or (cond2)`
  - Added: `_is_or_condition()` and `_split_or_conditions()` helpers

- [x] **P2-10** Enhance schematron parser for inequality ✅
  - File: `src/openxml_audit/codegen/schematron_loader.py`
  - Parse: `@attr != value`
  - Added: `SchematronType.ATTRIBUTE_NOT_EQUAL`

- [x] **P2-11** Enhance schematron parser for multi-AND ✅
  - File: `src/openxml_audit/codegen/schematron_loader.py`
  - Parse: `@a != X and @a != Y and @a != Z`
  - Added: `SchematronType.AND_CONDITION`

- [x] **P2-NEW** Enhance parser for conditional values ✅
  - Parse: `@attr and @other = value`
  - Added: `SchematronType.CONDITIONAL_VALUE`

- [x] **P2-NEW** Enhance parser for attribute presence ✅
  - Parse: `@attr`, `@a and @b`
  - Added: `SchematronType.ATTRIBUTES_PRESENT`

- [x] **P2-NEW** Support scientific notation and 'f' suffix ✅
  - Parse: `@attr >= -1.7E308`, `@attr >= 32767f`
  - Updated: `NUM` pattern in `_classify_rule()`

- [x] **P2-NEW** Support hyphenated attribute names ✅
  - Parse: `@emma:disjunction-type`
  - Updated: `ATTR` pattern in `_classify_rule()`

### 2.5 Verification

- [x] **P2-12** Test Phase 2 coverage ✅
  - 818/948 rules converted (86.3%)
  - All 96 tests pass
  - 0 UNKNOWN rules remaining

---

## Phase 3: Polish & Verify

- [ ] **P3-1** Coverage verification test
  - File: `tests/test_schematron_coverage.py`
  - Assert: All interpretable rules have corresponding constraints
  - Assert: No `SchematronType.UNKNOWN` remains

- [ ] **P3-2** Performance benchmark
  - Measure validation time before/after
  - Target: < 2x slower than current
  - Optimize hot paths if needed

- [ ] **P3-3** Error message review
  - Ensure all constraint errors are descriptive
  - Include attribute name, value, and constraint details
  - Add context path for debugging

---

## Summary

| Phase | Tasks | Rules Converted | Cumulative |
|-------|-------|-----------------|------------|
| 1 | 9 | 673 | 70% |
| 2 | 15 | 818 | 86.3% |
| 3 | 3 | - | - |

**Rules not yet convertible (130 remaining):**
- CROSS_PART_COUNT: 53 rules (requires package context)
- ELEMENT_REFERENCE: 23 rules (complex XPath lookups)
- RELATIONSHIP_TYPE: 37 rules (partial - needs relationship access)
- Other edge cases: 17 rules

**Total: 27 tasks**

---

## Execution Order

```
P1-1 → P1-2 → P1-3 → P1-4 → P1-5  (constraints, parallel)
         ↓
       P1-6 (bridge)
         ↓
       P1-7 (integration)
         ↓
       P1-8 (namespace resolution)
         ↓
       P1-9 (verify phase 1) → 70%
         ↓
P2-1 → P2-2 (compound, parallel)
         ↓
P2-3 → P2-4 → P2-5 → P2-6 (simple, parallel)
         ↓
P2-9 → P2-10 → P2-11 (parser, parallel)
         ↓
      P2-12 (verify phase 2) → 86.3%
         ↓
P2-7 → P2-8 (cross-part) → Future work
         ↓
P3-1 → P3-2 → P3-3 (polish)
```
