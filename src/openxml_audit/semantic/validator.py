"""Semantic validator for Open XML documents."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import ValidationError
from openxml_audit.semantic.attributes import SemanticConstraint
from openxml_audit.semantic.references import IdTracker, validate_unique_ids
from openxml_audit.namespaces import MC, OFFICE_DOC_RELATIONSHIPS
from openxml_audit.semantic.relationships import validate_part_relationships

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import OpenXmlPart


class SemanticValidator:
    """Validates XML documents against semantic constraints.

    This validator checks:
    - Attribute value relationships and dependencies
    - Relationship validity and targets
    - ID uniqueness and references
    - Cross-element constraints
    """

    def __init__(self, validate_unique_ids: bool = True) -> None:
        self._constraints: dict[str, list[SemanticConstraint]] = {}
        self._id_tracker = IdTracker()
        self._validate_unique_ids = validate_unique_ids

    def register_constraint(self, element_tag: str, constraint: SemanticConstraint) -> None:
        """Register a semantic constraint for an element type.

        Args:
            element_tag: The element tag (Clark notation) to apply constraint to.
            constraint: The constraint to apply.
        """
        if element_tag not in self._constraints:
            self._constraints[element_tag] = []
        self._constraints[element_tag].append(constraint)

    def validate_part(
        self, part: OpenXmlPart, context: ValidationContext
    ) -> list[ValidationError]:
        """Validate a part against semantic constraints.

        Args:
            part: The part to validate.
            context: The validation context.

        Returns:
            List of validation errors.
        """
        context.set_part(part)

        # Validate relationships
        validate_part_relationships(part, context)

        xml = part.xml
        if xml is None:
            return context.errors

        # Clear ID tracker for this part
        self._id_tracker.clear(part.uri)

        # Validate unique IDs when enabled
        if self._validate_unique_ids:
            validate_unique_ids(xml, "id", context, self._id_tracker, part.uri)

        # Validate element constraints
        self._validate_element(xml, context)

        return context.errors

    def _validate_element(self, element: etree._Element, context: ValidationContext) -> None:
        """Validate an element and its children recursively."""
        if context.should_stop:
            return

        with ElementContext(context, element):
            tag = element.tag

            # Validate relationship attributes for any OOXML element.
            self._validate_relationship_attributes(element, context)
            self._validate_mc_ignorable(element, context)

            # Apply registered constraints for this element type
            if tag in self._constraints:
                for constraint in self._constraints[tag]:
                    constraint.validate(element, context)

            # Recursively validate children
            for child in element:
                if isinstance(child.tag, str):
                    self._validate_element(child, context)

    def _validate_relationship_attributes(
        self, element: etree._Element, context: ValidationContext
    ) -> None:
        """Ensure relationship ID attributes reference existing relationships."""
        part = context.part
        if part is None:
            return
        for attr_name, value in element.attrib.items():
            if not attr_name.startswith(f"{{{OFFICE_DOC_RELATIONSHIPS}}}"):
                continue
            if not value:
                continue
            rel = part.relationships.get_by_id(value)
            if rel is None:
                local_attr = attr_name.split("}")[-1] if attr_name.startswith("{") else attr_name
                context.add_semantic_error(
                    f"Relationship '{value}' referenced by '{local_attr}' does not exist",
                    node=local_attr,
                )

    def _validate_mc_ignorable(
        self, element: etree._Element, context: ValidationContext
    ) -> None:
        ignorable = element.get(f"{{{MC}}}Ignorable")
        if ignorable is None:
            return
        prefixes = [prefix for prefix in ignorable.split() if prefix]
        if not prefixes:
            context.add_semantic_error(
                "Ignorable attribute is empty",
                node="Ignorable",
            )
            return
        nsmap = element.nsmap or {}
        for prefix in prefixes:
            if prefix not in nsmap or not nsmap.get(prefix):
                context.add_semantic_error(
                    f"Ignorable attribute contains undefined prefix '{prefix}'",
                    node="Ignorable",
                )


def create_pptx_semantic_validator(load_sdk_rules: bool = True) -> SemanticValidator:
    """Create a semantic validator with PPTX-specific constraints.

    Args:
        load_sdk_rules: If True, load all SDK schematron rules.

    Returns:
        A SemanticValidator configured for PPTX validation.
    """
    from openxml_audit.namespaces import DRAWINGML, PRESENTATIONML
    from openxml_audit.semantic.attributes import (
        AttributeMinMaxConstraint,
        AttributeValueLessEqualToAnother,
    )
    from openxml_audit.semantic.relationships import RelationshipExistConstraint

    validator = SemanticValidator(validate_unique_ids=True)

    # Load SDK schematron rules
    if load_sdk_rules:
        try:
            from openxml_audit.codegen.schematron_bridge import load_sdk_constraints

            for element_tag, constraint in load_sdk_constraints(app_filter="All"):
                validator.register_constraint(element_tag, constraint)
        except ImportError:
            pass  # SDK data not available

    # Relationship namespace
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    # Slide ID references (keep these as they check relationship existence, not just type)
    validator.register_constraint(
        f"{{{PRESENTATIONML}}}sldId",
        RelationshipExistConstraint(
            attribute="id",
            namespace=rel_ns,
        ),
    )

    # Slide master ID references
    validator.register_constraint(
        f"{{{PRESENTATIONML}}}sldMasterId",
        RelationshipExistConstraint(
            attribute="id",
            namespace=rel_ns,
        ),
    )

    # Notes master ID references
    validator.register_constraint(
        f"{{{PRESENTATIONML}}}notesMasterId",
        RelationshipExistConstraint(
            attribute="id",
            namespace=rel_ns,
        ),
    )

    return validator


def create_word_semantic_validator(load_sdk_rules: bool = True) -> SemanticValidator:
    """Create a semantic validator with Word-specific constraints."""
    from openxml_audit.namespaces import OFFICE_DOC_RELATIONSHIPS, REL_FONT, WORDPROCESSINGML
    from openxml_audit.semantic.attributes import (
        AttributeValueInSetConstraint,
        AttributeValuePatternConstraint,
    )
    from openxml_audit.semantic.relationships import RelationshipTypeConstraint

    validator = SemanticValidator(validate_unique_ids=False)

    if load_sdk_rules:
        try:
            from openxml_audit.codegen.schematron_bridge import load_sdk_constraints

            for element_tag, constraint in load_sdk_constraints(app_filter="Word"):
                validator.register_constraint(element_tag, constraint)
        except ImportError:
            pass

    zoom_values = ("none", "fullPage", "bestFit", "textFit")
    theme_colors = (
        "dark1",
        "light1",
        "dark2",
        "light2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hyperlink",
        "followedHyperlink",
    )

    validator.register_constraint(
        f"{{{WORDPROCESSINGML}}}zoom",
        AttributeValueInSetConstraint(
            attribute="val",
            namespace=WORDPROCESSINGML,
            allowed_values=zoom_values,
        ),
    )
    validator.register_constraint(
        f"{{{WORDPROCESSINGML}}}color",
        AttributeValueInSetConstraint(
            attribute="themeColor",
            namespace=WORDPROCESSINGML,
            allowed_values=theme_colors,
        ),
    )

    font_key_pattern = r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
    for tag in ("embedRegular", "embedBold", "embedItalic", "embedBoldItalic"):
        element_tag = f"{{{WORDPROCESSINGML}}}{tag}"
        validator.register_constraint(
            element_tag,
            AttributeValuePatternConstraint(
                attribute="fontKey",
                namespace=WORDPROCESSINGML,
                pattern=font_key_pattern,
            ),
        )
        validator.register_constraint(
            element_tag,
            RelationshipTypeConstraint(
                relationship_id_attribute="id",
                namespace=OFFICE_DOC_RELATIONSHIPS,
                expected_type=REL_FONT,
            ),
        )

    return validator


def create_spreadsheet_semantic_validator(load_sdk_rules: bool = True) -> SemanticValidator:
    """Create a semantic validator with Excel-specific constraints."""
    validator = SemanticValidator(validate_unique_ids=False)

    if load_sdk_rules:
        try:
            from openxml_audit.codegen.schematron_bridge import load_sdk_constraints

            for element_tag, constraint in load_sdk_constraints(app_filter="Excel"):
                validator.register_constraint(element_tag, constraint)
        except ImportError:
            pass

    return validator
