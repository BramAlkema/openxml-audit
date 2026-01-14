"""Schema validator for Open XML documents."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext, ValidationContext
from openxml_audit.errors import ValidationError
from openxml_audit.namespaces import MC
from openxml_audit.schema.constraints import get_constraint_for_tag as get_hardcoded_constraint
from openxml_audit.schema.particle import (
    CompositeParticle,
    ParticleType,
    get_validator,
)

# Try to import SDK constraint bridge
try:
    from openxml_audit.codegen.constraint_bridge import (
        get_element_constraint as get_sdk_constraint,
        get_element_constraint_for_element as get_sdk_constraint_for_element,
    )
    _HAS_SDK_CONSTRAINTS = True
except ImportError:
    _HAS_SDK_CONSTRAINTS = False
    get_sdk_constraint = None  # type: ignore
    get_sdk_constraint_for_element = None  # type: ignore


def get_constraint_for_tag(tag: str, element: etree._Element | None = None):
    """Get constraint for element tag, preferring SDK constraints.

    Args:
        tag: Element tag in Clark notation.

    Returns:
        ElementConstraint if found, None otherwise.
    """
    # Try SDK constraints first (more complete)
    if _HAS_SDK_CONSTRAINTS and get_sdk_constraint is not None:
        if element is not None and get_sdk_constraint_for_element is not None:
            constraint = get_sdk_constraint_for_element(tag, element)
        else:
            constraint = get_sdk_constraint(tag)
        if constraint is not None:
            return constraint

    # Fall back to hardcoded constraints
    return get_hardcoded_constraint(tag)

if TYPE_CHECKING:
    from openxml_audit.parts import OpenXmlPart


class SchemaValidator:
    """Validates XML documents against schema constraints.

    This validator checks:
    - Required attributes are present
    - Attribute values conform to type constraints
    - Child elements match content model (sequence, choice, all)
    """

    def __init__(self, validate_unknown_elements: bool = False):
        """Initialize the schema validator.

        Args:
            validate_unknown_elements: If True, report errors for elements
                                       without known constraints.
        """
        self._validate_unknown = validate_unknown_elements

    def validate_part(
        self, part: OpenXmlPart, context: ValidationContext
    ) -> list[ValidationError]:
        """Validate an XML part against schema constraints.

        Args:
            part: The part to validate.
            context: The validation context.

        Returns:
            List of validation errors.
        """
        context.set_part(part)

        xml = part.xml
        if xml is None:
            return context.errors

        self._validate_element(xml, context)

        return context.errors

    def _validate_element(self, element: etree._Element, context: ValidationContext) -> None:
        """Validate an element and its children recursively."""
        if context.should_stop:
            return

        with ElementContext(context, element):
            tag = element.tag

            # Get constraint for this element
            constraint = get_constraint_for_tag(tag, element)

            if constraint is not None:
                # Validate attributes
                self._validate_attributes(element, constraint, context)

                # Validate content model
                if constraint.content_model is not None:
                    self._validate_content_model(element, constraint.content_model, context)

            # Recursively validate children
            for child in self._get_validation_children(element):
                self._validate_element(child, context)

    def _validate_attributes(
        self,
        element: etree._Element,
        constraint: "ElementConstraint",  # type: ignore
        context: ValidationContext,
    ) -> None:
        """Validate element attributes."""
        # Check required attributes
        for attr_constraint in constraint.get_required_attributes():
            attr_name = attr_constraint.qualified_name
            if attr_name not in element.attrib:
                context.add_schema_error(
                    f"Required attribute '{attr_constraint.local_name}' is missing",
                    node=attr_constraint.local_name,
                )

        # Validate attribute values
        for attr_constraint in constraint.attributes:
            attr_name = attr_constraint.qualified_name
            if attr_name in element.attrib:
                value = element.attrib[attr_name]

                # Check fixed value
                if attr_constraint.fixed_value is not None:
                    if value != attr_constraint.fixed_value:
                        context.add_schema_error(
                            f"Attribute '{attr_constraint.local_name}' must have "
                            f"fixed value '{attr_constraint.fixed_value}', got '{value}'",
                            node=attr_constraint.local_name,
                        )

                # Type validation
                if attr_constraint.type_validator is not None:
                    result = attr_constraint.type_validator.validate(value, context)
                    if not result.is_valid:
                        context.add_schema_error(
                            f"Invalid value for attribute '{attr_constraint.local_name}': "
                            f"{result.error_message}",
                            node=attr_constraint.local_name,
                        )


    def _validate_content_model(
        self,
        element: etree._Element,
        content_model: "ParticleConstraint",  # type: ignore
        context: ValidationContext,
    ) -> None:
        """Validate element children against content model."""
        # Get non-comment children, with AlternateContent expanded
        children = self._get_validation_children(element)

        if isinstance(content_model, CompositeParticle):
            validator = get_validator(content_model.particle_type)
            if validator is not None:
                validator.validate(content_model, children, context)

    def _get_validation_children(self, element: etree._Element) -> list[etree._Element]:
        children: list[etree._Element] = []

        for child in element:
            if not isinstance(child.tag, str):
                continue
            if child.tag == f"{{{MC}}}AlternateContent":
                children.extend(self._resolve_alternate_content(child))
                continue
            children.append(child)

        return children

    def _resolve_alternate_content(self, alt: etree._Element) -> list[etree._Element]:
        ns = {"mc": MC}
        chosen = alt.find("mc:Fallback", ns)
        if chosen is None:
            chosen = alt.find("mc:Choice", ns)
        if chosen is None:
            return []
        return [c for c in chosen if isinstance(c.tag, str)]


# Import for type hints only
from openxml_audit.schema.constraints import ElementConstraint  # noqa: E402
from openxml_audit.schema.particle import ParticleConstraint  # noqa: E402
