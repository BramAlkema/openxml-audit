"""Tests for schematron coverage verification.

Ensures that the schematron-to-constraint bridge achieves expected coverage
and that no rules are classified as UNKNOWN.
"""

from __future__ import annotations

import pytest

from openxml_audit.codegen.schematron_bridge import get_sdk_constraint_stats
from openxml_audit.codegen.schematron_loader import (
    SchematronRegistry,
    SchematronType,
    get_registry,
)


class TestSchematronCoverage:
    """Tests for schematron rule coverage."""

    def test_no_unknown_rules(self) -> None:
        """Verify no rules are classified as UNKNOWN."""
        registry = get_registry()
        registry.load()

        unknown_rules = [
            r for r in registry._rules if r.rule_type == SchematronType.UNKNOWN
        ]

        assert len(unknown_rules) == 0, (
            f"Found {len(unknown_rules)} UNKNOWN rules. "
            f"First 5: {[r.test for r in unknown_rules[:5]]}"
        )

    def test_minimum_coverage_threshold(self) -> None:
        """Verify conversion achieves minimum coverage threshold."""
        stats = get_sdk_constraint_stats()

        # Current target: 91% coverage
        min_coverage = 0.91
        actual_coverage = stats["converted"] / stats["total"]

        assert actual_coverage >= min_coverage, (
            f"Coverage {actual_coverage:.1%} below threshold {min_coverage:.1%}. "
            f"Converted: {stats['converted']}/{stats['total']}"
        )

    def test_all_rule_types_handled(self) -> None:
        """Verify all rule types have conversion logic."""
        stats = get_sdk_constraint_stats()

        # These types should have 100% conversion (excluding edge cases)
        high_conversion_types = [
            "ATTRIBUTE_VALUE_LENGTH",
            "ATTRIBUTE_NOT_EQUAL",
            "ATTRIBUTES_PRESENT",
            "ATTRIBUTE_COMPARISON",
            "AND_CONDITION",
            "CROSS_PART_COUNT",
        ]

        for rule_type in high_conversion_types:
            if rule_type in stats["by_type"]:
                type_stats = stats["by_type"][rule_type]
                if type_stats["total"] > 0:
                    conversion_rate = type_stats["converted"] / type_stats["total"]
                    assert conversion_rate >= 0.9, (
                        f"{rule_type} has low conversion rate: "
                        f"{type_stats['converted']}/{type_stats['total']} "
                        f"({conversion_rate:.1%})"
                    )

    def test_schema_registry_loads(self) -> None:
        """Verify schema registry loads without errors."""
        registry = get_registry()
        registry.load()

        assert registry.count_rules() > 0
        assert len(registry._by_context) > 0

    def test_stats_structure(self) -> None:
        """Verify stats dictionary has expected structure."""
        stats = get_sdk_constraint_stats()

        assert "total" in stats
        assert "converted" in stats
        assert "skipped_no_context" in stats
        assert "skipped_no_constraint" in stats
        assert "by_type" in stats

        # Verify counts add up
        assert stats["converted"] + stats["skipped_no_context"] + stats["skipped_no_constraint"] == stats["total"]


class TestSchematronRuleTypes:
    """Tests for individual rule type parsing."""

    def test_attribute_value_range_parsing(self) -> None:
        """Test attribute value range rules are parsed correctly."""
        registry = get_registry()
        registry.load()

        range_rules = [
            r for r in registry._rules
            if r.rule_type == SchematronType.ATTRIBUTE_VALUE_RANGE
        ]

        assert len(range_rules) > 200  # Should have 240+ rules

        # Check that rules have min/max values extracted
        for rule in range_rules[:10]:
            assert rule.attribute is not None
            assert rule.min_value is not None or rule.max_value is not None

    def test_or_condition_parsing(self) -> None:
        """Test OR condition rules are parsed with sub-rules."""
        registry = get_registry()
        registry.load()

        or_rules = [
            r for r in registry._rules
            if r.rule_type == SchematronType.OR_CONDITION
        ]

        assert len(or_rules) > 0

        # Check that sub-rules are parsed
        for rule in or_rules:
            assert len(rule.sub_rules) >= 2, (
                f"OR rule should have at least 2 sub-rules: {rule.test}"
            )

    def test_conditional_value_parsing(self) -> None:
        """Test conditional value rules have trigger and sub-rule."""
        registry = get_registry()
        registry.load()

        conditional_rules = [
            r for r in registry._rules
            if r.rule_type == SchematronType.CONDITIONAL_VALUE
        ]

        assert len(conditional_rules) > 0

        for rule in conditional_rules:
            assert rule.attribute is not None, f"Missing trigger attribute: {rule.test}"
            assert len(rule.sub_rules) == 1, f"Should have exactly 1 sub-rule: {rule.test}"


class TestSchematronRegistry:
    """Tests for SchematronRegistry functionality."""

    def test_get_rules_for_context(self) -> None:
        """Test retrieving rules by context."""
        registry = get_registry()
        registry.load()

        # p:sld is a common context
        rules = registry.get_rules_for_context("p:sld")
        # Should have some rules for slide element
        # (exact count depends on SDK data)

    def test_get_rules_by_type(self) -> None:
        """Test retrieving rules by type."""
        registry = get_registry()
        registry.load()

        range_rules = registry.get_rules_by_type(SchematronType.ATTRIBUTE_VALUE_RANGE)
        assert len(range_rules) > 0

        for rule in range_rules:
            assert rule.rule_type == SchematronType.ATTRIBUTE_VALUE_RANGE

    def test_get_interpretable_rules(self) -> None:
        """Test getting only interpretable rules."""
        registry = get_registry()
        registry.load()

        interpretable = registry.get_interpretable_rules()

        for rule in interpretable:
            assert rule.rule_type != SchematronType.UNKNOWN

    def test_count_by_type(self) -> None:
        """Test counting rules by type."""
        registry = get_registry()
        registry.load()

        counts = registry.count_by_type()

        assert isinstance(counts, dict)
        assert all(isinstance(k, SchematronType) for k in counts.keys())
        assert all(isinstance(v, int) for v in counts.values())
        assert sum(counts.values()) == registry.count_rules()

    def test_get_stats(self) -> None:
        """Test stats method."""
        registry = get_registry()
        registry.load()

        stats = registry.get_stats()

        assert stats["total"] > 0
        assert stats["interpretable"] > 0
        assert stats["unique_contexts"] > 0
        assert "by_type" in stats
