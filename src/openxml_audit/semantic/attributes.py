"""Semantic attribute constraints for Open XML validation.

These constraints go beyond schema validation to enforce business rules
and relationships between attribute values.
"""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext


class SemanticConstraint(ABC):
    """Base class for semantic constraints."""

    @abstractmethod
    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        """Validate the constraint against an element.

        Args:
            element: The element to validate.
            context: The validation context.

        Returns:
            True if valid, False otherwise.
        """
        pass


@dataclass
class AttributeMinMaxConstraint(SemanticConstraint):
    """Validates that an attribute value is within a min/max range.

    Unlike schema type constraints, this can reference values
    from other attributes dynamically.
    """

    attribute: str
    min_value: int | float | None = None
    max_value: int | float | None = None
    min_inclusive: bool = True
    max_inclusive: bool = True
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True  # Not present, nothing to validate

        try:
            value = float(element.attrib[attr_name])
        except ValueError:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' must be numeric",
                node=self.attribute,
            )
            return False

        if self.min_value is not None:
            if self.min_inclusive and value < self.min_value:
                context.add_semantic_error(
                    f"Attribute '{self.attribute}' value {value} is less than minimum {self.min_value}",
                    node=self.attribute,
                )
                return False
            if not self.min_inclusive and value <= self.min_value:
                context.add_semantic_error(
                    f"Attribute '{self.attribute}' value {value} must be greater than {self.min_value}",
                    node=self.attribute,
                )
                return False

        if self.max_value is not None:
            if self.max_inclusive and value > self.max_value:
                context.add_semantic_error(
                    f"Attribute '{self.attribute}' value {value} exceeds maximum {self.max_value}",
                    node=self.attribute,
                )
                return False
            if not self.max_inclusive and value >= self.max_value:
                context.add_semantic_error(
                    f"Attribute '{self.attribute}' value {value} must be less than {self.max_value}",
                    node=self.attribute,
                )
                return False

        return True


@dataclass
class AttributeValuePatternConstraint(SemanticConstraint):
    """Validates that an attribute value matches a pattern."""

    attribute: str
    pattern: str
    namespace: str | None = None
    error_message: str | None = None

    def __post_init__(self) -> None:
        self._compiled = re.compile(self.pattern)

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        value = element.attrib[attr_name]
        if not self._compiled.match(value):
            msg = self.error_message or f"Attribute '{self.attribute}' value '{value}' does not match required pattern"
            context.add_semantic_error(msg, node=self.attribute)
            return False

        return True


@dataclass
class AttributeMutualExclusive(SemanticConstraint):
    """Validates that only one of a set of attributes is present."""

    attributes: list[str]
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        present = []
        for attr in self.attributes:
            attr_name = f"{{{self.namespace}}}{attr}" if self.namespace else attr
            if attr_name in element.attrib:
                present.append(attr)

        if len(present) > 1:
            context.add_semantic_error(
                f"Attributes {present} are mutually exclusive - only one can be present",
            )
            return False

        return True


@dataclass
class AttributeRequiredConditionToValue(SemanticConstraint):
    """Validates that an attribute is required when another has a specific value."""

    required_attribute: str
    condition_attribute: str
    condition_value: str
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        cond_name = f"{{{self.namespace}}}{self.condition_attribute}" if self.namespace else self.condition_attribute
        req_name = f"{{{self.namespace}}}{self.required_attribute}" if self.namespace else self.required_attribute

        if cond_name not in element.attrib:
            return True

        if element.attrib[cond_name] == self.condition_value:
            if req_name not in element.attrib:
                context.add_semantic_error(
                    f"Attribute '{self.required_attribute}' is required when "
                    f"'{self.condition_attribute}' is '{self.condition_value}'",
                    node=self.required_attribute,
                )
                return False

        return True


@dataclass
class AttributeValueLengthConstraint(SemanticConstraint):
    """Validates attribute value length."""

    attribute: str
    min_length: int | None = None
    max_length: int | None = None
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        value = element.attrib[attr_name]
        length = len(value)

        if self.min_length is not None and length < self.min_length:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' length {length} is less than minimum {self.min_length}",
                node=self.attribute,
            )
            return False

        if self.max_length is not None and length > self.max_length:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' length {length} exceeds maximum {self.max_length}",
                node=self.attribute,
            )
            return False

        return True


@dataclass
class AttributeValueInSetConstraint(SemanticConstraint):
    """Validates that an attribute value is one of a set of allowed values."""

    attribute: str
    allowed_values: set[str] | tuple[str, ...]
    namespace: str | None = None
    error_message: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        value = element.attrib[attr_name]
        if value not in self.allowed_values:
            msg = (
                self.error_message
                or f"Attribute '{self.attribute}' has invalid value '{value}'"
            )
            context.add_semantic_error(msg, node=self.attribute)
            return False

        return True


@dataclass
class AttributeValueRangeConstraint(SemanticConstraint):
    """Validates that an attribute value references another attribute as a range."""

    attribute: str
    min_attribute: str | None = None
    max_attribute: str | None = None
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        try:
            value = float(element.attrib[attr_name])
        except ValueError:
            return True  # Let schema validation handle type errors

        # Check against min attribute
        if self.min_attribute:
            min_name = f"{{{self.namespace}}}{self.min_attribute}" if self.namespace else self.min_attribute
            if min_name in element.attrib:
                try:
                    min_val = float(element.attrib[min_name])
                    if value < min_val:
                        context.add_semantic_error(
                            f"Attribute '{self.attribute}' value {value} is less than "
                            f"'{self.min_attribute}' value {min_val}",
                            node=self.attribute,
                        )
                        return False
                except ValueError:
                    pass

        # Check against max attribute
        if self.max_attribute:
            max_name = f"{{{self.namespace}}}{self.max_attribute}" if self.namespace else self.max_attribute
            if max_name in element.attrib:
                try:
                    max_val = float(element.attrib[max_name])
                    if value > max_val:
                        context.add_semantic_error(
                            f"Attribute '{self.attribute}' value {value} exceeds "
                            f"'{self.max_attribute}' value {max_val}",
                            node=self.attribute,
                        )
                        return False
                except ValueError:
                    pass

        return True


@dataclass
class AttributeValueSetConstraint(SemanticConstraint):
    """Validates that an attribute value is in a set of allowed values."""

    attribute: str
    allowed_values: set[str]
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        value = element.attrib[attr_name]
        if value not in self.allowed_values:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' value '{value}' is not in allowed set: "
                f"{sorted(self.allowed_values)}",
                node=self.attribute,
            )
            return False

        return True


@dataclass
class AttributeValueLessEqualToAnother(SemanticConstraint):
    """Validates that one attribute value is <= another."""

    attribute: str
    other_attribute: str
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute
        other_name = f"{{{self.namespace}}}{self.other_attribute}" if self.namespace else self.other_attribute

        if attr_name not in element.attrib or other_name not in element.attrib:
            return True

        try:
            value = float(element.attrib[attr_name])
            other_value = float(element.attrib[other_name])
        except ValueError:
            return True

        if value > other_value:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' value {value} must be <= "
                f"'{self.other_attribute}' value {other_value}",
                node=self.attribute,
            )
            return False

        return True
