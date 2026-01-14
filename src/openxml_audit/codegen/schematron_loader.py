"""Load and interpret Open XML SDK schematron rules at runtime.

Schematron rules provide semantic validation beyond schema constraints,
such as attribute value ranges, uniqueness, and cross-references.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from lxml import etree

# Path to SDK data files
DATA_DIR = Path(__file__).parent.parent.parent.parent / "data" / "openxml"


class SchematronType(Enum):
    """Types of schematron rules we can interpret."""

    ATTRIBUTE_VALUE_RANGE = auto()  # @attr >= N and @attr <= M
    ATTRIBUTE_VALUE_LENGTH = auto()  # string-length(@attr) <= N
    ATTRIBUTE_VALUE_PATTERN = auto()  # matches(@attr, 'pattern')
    UNIQUE_ATTRIBUTE = auto()  # count(distinct-values(...)) = count(...)
    RELATIONSHIP_TYPE = auto()  # document(rels)//r:Relationship[...]
    ELEMENT_REFERENCE = auto()  # Index-of(document(...), @id)
    ATTRIBUTE_NOT_EQUAL = auto()  # @attr != value
    ATTRIBUTE_EQUALS = auto()  # @attr = value
    ATTRIBUTE_COMPARISON = auto()  # @attr < @other_attr
    OR_CONDITION = auto()  # (cond1) or (cond2)
    AND_CONDITION = auto()  # cond1 and cond2 (multi-attribute)
    ATTRIBUTES_PRESENT = auto()  # @a and @b (both present)
    CROSS_PART_COUNT = auto()  # @attr < count(document('Part:...')//...)
    CONDITIONAL_VALUE = auto()  # @attr and @other = value (if attr present, other must equal)
    UNKNOWN = auto()  # Cannot be auto-interpreted


@dataclass
class ParsedSchematron:
    """A parsed schematron rule."""

    context: str  # XPath context (e.g., "p:sld")
    test: str  # XPath test expression
    app: str  # Application scope: "All", "PowerPoint", "Word", "Excel"
    rule_type: SchematronType = SchematronType.UNKNOWN

    # Extracted parameters (depends on rule_type)
    attribute: str | None = None
    min_value: float | None = None
    max_value: float | None = None
    min_length: int | None = None
    max_length: int | None = None
    pattern: str | None = None
    relationship_type: str | None = None
    # For equality/comparison
    expected_value: str | None = None
    forbidden_value: str | None = None
    other_attribute: str | None = None
    comparison_operator: str | None = None
    # For compound rules
    sub_rules: list["ParsedSchematron"] = field(default_factory=list)
    # For attributes present (@a and @b)
    required_attributes: list[str] = field(default_factory=list)
    # For cross-part count validation
    part_path: str | None = None
    element_xpath: str | None = None
    count_offset: int = 0

    @property
    def context_prefix(self) -> str | None:
        """Get the namespace prefix from context."""
        if ":" in self.context:
            return self.context.split(":")[0]
        return None

    @property
    def context_local_name(self) -> str:
        """Get the local name from context."""
        if ":" in self.context:
            return self.context.split(":", 1)[1]
        return self.context


def parse_schematron(data: dict[str, Any]) -> ParsedSchematron:
    """Parse a schematron rule and extract its type and parameters."""
    context = data["Context"]
    test = data["Test"]
    app = data.get("App", "All")

    rule = ParsedSchematron(context=context, test=test, app=app)

    # Try to classify and extract parameters
    _classify_rule(rule)

    return rule


def _is_or_condition(test: str) -> bool:
    """Check if test contains a top-level 'or' (not inside parentheses)."""
    depth = 0
    i = 0
    while i < len(test):
        if test[i] == "(":
            depth += 1
        elif test[i] == ")":
            depth -= 1
        elif depth == 0 and test[i : i + 4] == " or ":
            return True
        i += 1
    return False


def _split_or_conditions(test: str) -> list[str]:
    """Split test by top-level 'or' (not inside parentheses)."""
    parts: list[str] = []
    depth = 0
    current = ""
    i = 0
    while i < len(test):
        if test[i] == "(":
            depth += 1
            current += test[i]
        elif test[i] == ")":
            depth -= 1
            current += test[i]
        elif depth == 0 and test[i : i + 4] == " or ":
            parts.append(current.strip())
            current = ""
            i += 3  # Skip " or"
        else:
            current += test[i]
        i += 1
    if current.strip():
        parts.append(current.strip())
    return parts


def _classify_rule(rule: ParsedSchematron) -> None:
    """Classify a schematron rule and extract parameters."""
    test = rule.test

    # Pattern for attribute names (including prefixed and hyphenated names)
    # Matches: attr, prefix:attr, prefix:attr-name
    ATTR = r'[\w:-]+'

    # Pattern for numbers including scientific notation (e.g., -1.7E308) and f suffix
    NUM = r'[\d.eE+-]+f?'

    # Pattern: @attr >= N and @attr <= M (attribute value range)
    range_match = re.match(
        rf'@(\w+:?\w*)\s*>=?\s*({NUM})\s+and\s+@\1\s*<=?\s*({NUM})',
        test
    )
    if range_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_RANGE
        rule.attribute = range_match.group(1)
        rule.min_value = float(range_match.group(2).rstrip('f'))
        rule.max_value = float(range_match.group(3).rstrip('f'))
        return

    # Pattern: @attr <= N (single upper bound)
    upper_match = re.match(rf'@(\w+:?\w*)\s*<=?\s*({NUM})$', test)
    if upper_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_RANGE
        rule.attribute = upper_match.group(1)
        rule.max_value = float(upper_match.group(2).rstrip('f'))
        return

    # Pattern: @attr >= N (single lower bound)
    lower_match = re.match(rf'@(\w+:?\w*)\s*>=?\s*({NUM})$', test)
    if lower_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_RANGE
        rule.attribute = lower_match.group(1)
        rule.min_value = float(lower_match.group(2).rstrip('f'))
        return

    # Pattern: string-length(@attr) <= N or >= N and <= M
    strlen_match = re.match(
        r'string-length\(@(\w+:?\w*)\)\s*>=?\s*(\d+)\s+and\s+string-length\(@\1\)\s*<=?\s*(\d+)',
        test
    )
    if strlen_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_LENGTH
        rule.attribute = strlen_match.group(1)
        rule.min_length = int(strlen_match.group(2))
        rule.max_length = int(strlen_match.group(3))
        return

    strlen_max_match = re.match(r'string-length\(@(\w+:?\w*)\)\s*<=?\s*(\d+)', test)
    if strlen_max_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_LENGTH
        rule.attribute = strlen_max_match.group(1)
        rule.max_length = int(strlen_max_match.group(2))
        return

    # Pattern: string-length(@attr) >= N (min-only length)
    strlen_min_match = re.match(r'string-length\(@(\w+:?\w*)\)\s*>=?\s*(\d+)$', test)
    if strlen_min_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_LENGTH
        rule.attribute = strlen_min_match.group(1)
        rule.min_length = int(strlen_min_match.group(2))
        return

    # Pattern: matches(@attr, "pattern")
    pattern_match = re.match(r'matches\(@(\w+:?\w*),\s*["\'](.+?)["\']\)', test)
    if pattern_match:
        rule.rule_type = SchematronType.ATTRIBUTE_VALUE_PATTERN
        rule.attribute = pattern_match.group(1)
        rule.pattern = pattern_match.group(2)
        return

    # Pattern: count(distinct-values(...)) = count(...) (uniqueness)
    if "count(distinct-values(" in test and "= count(" in test:
        rule.rule_type = SchematronType.UNIQUE_ATTRIBUTE
        # Extract attribute from the expression
        attr_match = re.search(r'/@(\w+:?\w*)\)', test)
        if attr_match:
            rule.attribute = attr_match.group(1)
        return

    # Pattern: document(rels)//r:Relationship[...] (relationship check)
    if "document(rels)" in test and "r:Relationship" in test:
        rule.rule_type = SchematronType.RELATIONSHIP_TYPE
        # Extract relationship type
        type_match = re.search(r"@Type\s*=\s*['\"](.+?)['\"]", test)
        if type_match:
            rule.relationship_type = type_match.group(1)
        # Extract attribute being checked
        attr_match = re.search(r"@Id\s*=\s*current\(\)/@(\w+:?\w*)", test)
        if attr_match:
            rule.attribute = attr_match.group(1)
        return

    # Pattern: Index-of(document(...), @id) (reference check)
    if "Index-of(document(" in test or "index-of(document(" in test.lower():
        rule.rule_type = SchematronType.ELEMENT_REFERENCE
        return

    # Pattern: @attr != value (attribute not equal)
    not_equal_match = re.match(r"@(\w+:?\w*)\s*!=\s*['\"]?([^'\"]+)['\"]?$", test)
    if not_equal_match:
        rule.rule_type = SchematronType.ATTRIBUTE_NOT_EQUAL
        rule.attribute = not_equal_match.group(1)
        rule.forbidden_value = not_equal_match.group(2).strip()
        return

    # Pattern: @attr = value (attribute equals)
    # Use ATTR to match hyphenated names like @emma:disjunction-type
    equals_match = re.match(rf"@({ATTR})\s*=\s*['\"]?([^'\"]+)['\"]?$", test)
    if equals_match:
        rule.rule_type = SchematronType.ATTRIBUTE_EQUALS
        rule.attribute = equals_match.group(1)
        rule.expected_value = equals_match.group(2).strip()
        return

    # Pattern: @attr < @other or @attr <= @other (attribute comparison)
    comparison_match = re.match(
        r"@(\w+:?\w*)\s*(<=?|>=?)\s*@(\w+:?\w*)$", test
    )
    if comparison_match:
        rule.rule_type = SchematronType.ATTRIBUTE_COMPARISON
        rule.attribute = comparison_match.group(1)
        rule.comparison_operator = comparison_match.group(2)
        rule.other_attribute = comparison_match.group(3)
        return

    # Pattern: (cond1) or (cond2) (OR condition)
    # Look for " or " as a top-level connector (not inside parentheses)
    if _is_or_condition(test):
        rule.rule_type = SchematronType.OR_CONDITION
        # Parse sub-conditions
        sub_tests = _split_or_conditions(test)
        for sub_test in sub_tests:
            sub_rule = ParsedSchematron(
                context=rule.context, test=sub_test.strip(), app=rule.app
            )
            _classify_rule(sub_rule)
            rule.sub_rules.append(sub_rule)
        return

    # Pattern: @a != X and @a != Y (multi-AND with same attribute)
    if " and " in test and "@" in test and "!=" in test:
        and_parts = re.split(r"\s+and\s+", test)
        if len(and_parts) >= 2:
            # Check if all parts are inequality checks
            all_not_equal = True
            for part in and_parts:
                if not re.match(r"@\w+:?\w*\s*!=", part.strip()):
                    all_not_equal = False
                    break
            if all_not_equal:
                rule.rule_type = SchematronType.AND_CONDITION
                for part in and_parts:
                    sub_rule = ParsedSchematron(
                        context=rule.context, test=part.strip(), app=rule.app
                    )
                    _classify_rule(sub_rule)
                    rule.sub_rules.append(sub_rule)
                return

    # Pattern: @attr (single attribute must be present)
    single_attr = re.match(r'^@(\w+:?\w*)$', test)
    if single_attr:
        rule.rule_type = SchematronType.ATTRIBUTES_PRESENT
        rule.required_attributes = [single_attr.group(1)]
        return

    # Pattern: @a and @b (attributes must both be present)
    # Simple case: only @attr references connected by 'and'
    attrs_present = re.match(r'^(@\w+:?\w*)(\s+and\s+@\w+:?\w*)+$', test)
    if attrs_present:
        rule.rule_type = SchematronType.ATTRIBUTES_PRESENT
        rule.required_attributes = re.findall(r'@(\w+:?\w*)', test)
        return

    # Pattern: @attr and <condition> (conditional - if attr present, condition must hold)
    # This catches: @a and @b = value, @a and @b != value, @a and (@b = x or @b = y)
    conditional_match = re.match(r'^@(\w+:?\w*)\s+and\s+(.+)$', test)
    if conditional_match:
        rule.rule_type = SchematronType.CONDITIONAL_VALUE
        rule.attribute = conditional_match.group(1)
        # Parse the condition as a sub-rule
        condition_test = conditional_match.group(2)
        sub_rule = ParsedSchematron(
            context=rule.context, test=condition_test.strip(), app=rule.app
        )
        _classify_rule(sub_rule)
        rule.sub_rules.append(sub_rule)
        return

    # Pattern: @attr < count(document('Part:...')//xpath) + N (cross-part count)
    cross_part = re.match(
        r"@(\w+:?\w*)\s*<\s*count\(document\(['\"]Part:([^'\"]+)['\"]\)//([^)]+)\)\s*\+\s*(\d+)",
        test
    )
    if cross_part:
        rule.rule_type = SchematronType.CROSS_PART_COUNT
        rule.attribute = cross_part.group(1)
        rule.part_path = cross_part.group(2)
        rule.element_xpath = cross_part.group(3)
        rule.count_offset = int(cross_part.group(4))
        return

    # Could not classify
    rule.rule_type = SchematronType.UNKNOWN


class SchematronRegistry:
    """Registry of schematron rules indexed by context element."""

    def __init__(self) -> None:
        self._rules: list[ParsedSchematron] = []
        self._by_context: dict[str, list[ParsedSchematron]] = {}
        self._by_prefix: dict[str, list[ParsedSchematron]] = {}
        self._loaded = False

    def load(self) -> None:
        """Load schematrons from disk."""
        if self._loaded:
            return

        schematron_file = DATA_DIR / "schematrons.json"
        if not schematron_file.exists():
            self._loaded = True
            return

        with open(schematron_file) as f:
            data = json.load(f)

        for item in data:
            rule = parse_schematron(item)
            self._rules.append(rule)

            # Index by exact context
            if rule.context not in self._by_context:
                self._by_context[rule.context] = []
            self._by_context[rule.context].append(rule)

            # Index by prefix for namespace-based lookup
            if rule.context_prefix:
                if rule.context_prefix not in self._by_prefix:
                    self._by_prefix[rule.context_prefix] = []
                self._by_prefix[rule.context_prefix].append(rule)

        self._loaded = True

    def get_rules_for_context(self, context: str) -> list[ParsedSchematron]:
        """Get all rules for a specific context (e.g., 'p:sld')."""
        self.load()
        return self._by_context.get(context, [])

    def get_rules_for_element(
        self,
        prefix: str,
        local_name: str,
        app: str = "All"
    ) -> list[ParsedSchematron]:
        """Get applicable rules for an element."""
        self.load()
        context = f"{prefix}:{local_name}" if prefix else local_name
        rules = self._by_context.get(context, [])

        # Filter by app
        if app != "All":
            rules = [r for r in rules if r.app in ("All", app)]

        return rules

    def get_rules_by_type(self, rule_type: SchematronType) -> list[ParsedSchematron]:
        """Get all rules of a specific type."""
        self.load()
        return [r for r in self._rules if r.rule_type == rule_type]

    def get_interpretable_rules(self) -> list[ParsedSchematron]:
        """Get rules that can be automatically interpreted."""
        self.load()
        return [r for r in self._rules if r.rule_type != SchematronType.UNKNOWN]

    def count_rules(self) -> int:
        """Count total number of rules."""
        self.load()
        return len(self._rules)

    def count_by_type(self) -> dict[SchematronType, int]:
        """Count rules by type."""
        self.load()
        counts: dict[SchematronType, int] = {}
        for rule in self._rules:
            counts[rule.rule_type] = counts.get(rule.rule_type, 0) + 1
        return counts

    def get_stats(self) -> dict[str, Any]:
        """Get statistics about loaded schematrons."""
        self.load()
        counts = self.count_by_type()
        interpretable = sum(
            c for t, c in counts.items() if t != SchematronType.UNKNOWN
        )
        return {
            "total": len(self._rules),
            "interpretable": interpretable,
            "by_type": {t.name: c for t, c in counts.items()},
            "unique_contexts": len(self._by_context),
        }


# Global registry instance
_registry: SchematronRegistry | None = None


def get_registry() -> SchematronRegistry:
    """Get the global schematron registry."""
    global _registry
    if _registry is None:
        _registry = SchematronRegistry()
    return _registry
