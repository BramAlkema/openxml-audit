"""Equality and inequality constraints for attribute values."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.semantic.attributes import SemanticConstraint

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext


@dataclass
class AttributeEqualsConstraint(SemanticConstraint):
    """Validates that an attribute equals a specific value.

    Used for schematron rules like:
        @x:spt = 19
        @x:bx = false
    """

    attribute: str
    expected_value: str
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = (
            f"{{{self.namespace}}}{self.attribute}"
            if self.namespace
            else self.attribute
        )

        if attr_name not in element.attrib:
            return True  # Attribute not present, nothing to validate

        value = element.attrib[attr_name]
        if value != self.expected_value:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' must equal '{self.expected_value}', "
                f"got '{value}'",
                node=self.attribute,
            )
            return False

        return True


@dataclass
class AttributeNotEqualConstraint(SemanticConstraint):
    """Validates that an attribute does not equal a specific value.

    Used for schematron rules like:
        @x:guid != 00000000-0000-0000-0000-000000000000
        @x:axis != axisValues
    """

    attribute: str
    forbidden_value: str
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = (
            f"{{{self.namespace}}}{self.attribute}"
            if self.namespace
            else self.attribute
        )

        if attr_name not in element.attrib:
            return True  # Attribute not present, nothing to validate

        value = element.attrib[attr_name]
        if value == self.forbidden_value:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' must not equal '{self.forbidden_value}'",
                node=self.attribute,
            )
            return False

        return True


@dataclass
class AttributesPresentConstraint(SemanticConstraint):
    """Validates that certain attributes are present together.

    Used for schematron rules like:
        @x:l and @x:s  (both must be present)
    """

    attributes: list[str]
    namespace: str | None = None
    all_required: bool = True  # If True, all must be present; if False, at least one

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        present = []
        missing = []

        for attr in self.attributes:
            attr_name = (
                f"{{{self.namespace}}}{attr}"
                if self.namespace
                else attr
            )
            if attr_name in element.attrib:
                present.append(attr)
            else:
                missing.append(attr)

        if self.all_required:
            # All attributes must be present if any are present
            if present and missing:
                context.add_semantic_error(
                    f"Attributes {missing} are required when {present} are present",
                )
                return False
        else:
            # At least one must be present
            if not present:
                context.add_semantic_error(
                    f"At least one of {self.attributes} must be present",
                )
                return False

        return True


@dataclass
class AttributeComparisonConstraint(SemanticConstraint):
    """Validates comparison between two attributes.

    Used for schematron rules like:
        @x:sb < @x:eb
    """

    attribute: str
    other_attribute: str
    operator: str  # "<", "<=", ">", ">=", "=", "!="
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = (
            f"{{{self.namespace}}}{self.attribute}"
            if self.namespace
            else self.attribute
        )
        other_name = (
            f"{{{self.namespace}}}{self.other_attribute}"
            if self.namespace
            else self.other_attribute
        )

        if attr_name not in element.attrib or other_name not in element.attrib:
            return True  # Attributes not present

        try:
            value = float(element.attrib[attr_name])
            other_value = float(element.attrib[other_name])
        except ValueError:
            return True  # Non-numeric, let schema validation handle

        result = self._compare(value, other_value)
        if not result:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' ({value}) must be {self.operator} "
                f"'{self.other_attribute}' ({other_value})",
                node=self.attribute,
            )
            return False

        return True

    def _compare(self, a: float, b: float) -> bool:
        match self.operator:
            case "<":
                return a < b
            case "<=":
                return a <= b
            case ">":
                return a > b
            case ">=":
                return a >= b
            case "=":
                return a == b
            case "!=":
                return a != b
            case _:
                return True
