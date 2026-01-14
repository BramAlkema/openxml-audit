"""Compound constraints (OR, AND) for combining multiple constraints."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.semantic.attributes import SemanticConstraint

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext


@dataclass
class OrConstraint(SemanticConstraint):
    """Validates that at least one sub-constraint passes.

    Used for schematron rules like:
        (@x:operator and @x:type = cells) or @x:type != cells
    """

    constraints: list[SemanticConstraint] = field(default_factory=list)
    error_message: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        if not self.constraints:
            return True

        # Check if any constraint passes
        for constraint in self.constraints:
            # Temporarily suppress errors - we only want to report if ALL fail
            original_errors = len(context.errors)
            result = constraint.validate(element, context)

            if result:
                # One passed - remove any errors added by failed constraints
                while len(context.errors) > original_errors:
                    context.errors.pop()
                return True

            # Remove errors from this failed attempt
            while len(context.errors) > original_errors:
                context.errors.pop()

        # All constraints failed
        if self.error_message:
            context.add_semantic_error(self.error_message)
        else:
            context.add_semantic_error(
                "None of the alternative conditions are satisfied"
            )
        return False


@dataclass
class AndConstraint(SemanticConstraint):
    """Validates that all sub-constraints pass.

    Used for schematron rules like:
        @x:maxValue != NaN and @x:maxValue != INF and @x:maxValue != -INF
    """

    constraints: list[SemanticConstraint] = field(default_factory=list)

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        if not self.constraints:
            return True

        # All constraints must pass
        all_passed = True
        for constraint in self.constraints:
            if not constraint.validate(element, context):
                all_passed = False
                # Continue checking to report all errors

        return all_passed


@dataclass
class ConditionalConstraint(SemanticConstraint):
    """Validates a condition only if an attribute is present.

    Used for schematron rules like:
        @x:sourceRef and @x:sourceType != range
    This means: IF sourceRef is present, THEN sourceType must != range
    """

    trigger_attribute: str
    constraint: SemanticConstraint
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = (
            f"{{{self.namespace}}}{self.trigger_attribute}"
            if self.namespace
            else self.trigger_attribute
        )

        # If trigger attribute is not present, constraint is satisfied
        if attr_name not in element.attrib:
            return True

        # Trigger attribute present - validate the condition
        return self.constraint.validate(element, context)
