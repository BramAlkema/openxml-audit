"""Semantic relationship constraints for Open XML validation.

These constraints validate that relationships between parts are valid
and that relationship references within the XML are correct.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.semantic.attributes import SemanticConstraint

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import OpenXmlPart


@dataclass
class RelationshipExistConstraint(SemanticConstraint):
    """Validates that an attribute referencing a relationship ID points to an existing relationship."""

    attribute: str
    namespace: str | None = None
    relationship_type: str | None = None  # Optional: require specific relationship type

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        rel_id = element.attrib[attr_name]

        # Get the current part's relationships
        part = context.part
        if part is None:
            return True

        rel = part.relationships.get_by_id(rel_id)
        if rel is None:
            context.add_semantic_error(
                f"Relationship '{rel_id}' referenced by '{self.attribute}' does not exist",
                node=self.attribute,
            )
            return False

        # Check relationship type if specified
        if self.relationship_type and rel.type != self.relationship_type:
            context.add_semantic_error(
                f"Relationship '{rel_id}' has type '{rel.type}' but expected '{self.relationship_type}'",
                node=self.attribute,
            )
            return False

        return True


@dataclass
class RelationshipTypeConstraint(SemanticConstraint):
    """Validates that a relationship has the correct type."""

    relationship_id_attribute: str
    expected_type: str
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = (
            f"{{{self.namespace}}}{self.relationship_id_attribute}"
            if self.namespace
            else self.relationship_id_attribute
        )

        if attr_name not in element.attrib:
            return True

        rel_id = element.attrib[attr_name]

        part = context.part
        if part is None:
            return True

        rel = part.relationships.get_by_id(rel_id)
        if rel is None:
            return True  # Let RelationshipExistConstraint handle missing rels

        if rel.type != self.expected_type:
            context.add_semantic_error(
                f"Relationship '{rel_id}' should be type '{self.expected_type}' "
                f"but is '{rel.type}'",
                node=self.relationship_id_attribute,
            )
            return False

        return True


@dataclass
class RelationshipTargetExistsConstraint(SemanticConstraint):
    """Validates that relationship targets point to existing parts."""

    relationship_id_attribute: str
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = (
            f"{{{self.namespace}}}{self.relationship_id_attribute}"
            if self.namespace
            else self.relationship_id_attribute
        )

        if attr_name not in element.attrib:
            return True

        rel_id = element.attrib[attr_name]

        part = context.part
        if part is None or context.package is None:
            return True

        rel = part.relationships.get_by_id(rel_id)
        if rel is None:
            return True  # Let RelationshipExistConstraint handle this

        if rel.is_external:
            return True  # External relationships don't need part validation

        target = rel.resolve_target(part.uri)
        if not context.package.has_part(target):
            context.add_semantic_error(
                f"Relationship '{rel_id}' target '{target}' does not exist in package",
                node=self.relationship_id_attribute,
            )
            return False

        return True


def validate_part_relationships(part: "OpenXmlPart", context: "ValidationContext") -> bool:
    """Validate all relationships for a part.

    Checks:
    - All internal relationship targets exist
    - No duplicate relationship IDs

    Args:
        part: The part to validate.
        context: The validation context.

    Returns:
        True if all relationships are valid.
    """
    if context.package is None:
        return True

    valid = True
    seen_ids: set[str] = set()

    for rel in part.relationships:
        # Check for duplicate IDs
        if rel.id in seen_ids:
            context.add_semantic_error(
                f"Duplicate relationship ID: '{rel.id}'",
                node=rel.id,
            )
            valid = False
        seen_ids.add(rel.id)

        # Check internal targets exist
        if not rel.is_external:
            target = rel.resolve_target(part.uri)
            if not context.package.has_part(target):
                context.add_semantic_error(
                    f"Relationship '{rel.id}' target not found: '{target}'",
                    node=rel.id,
                )
                valid = False

    return valid


# Common relationship namespace
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Pre-built constraints for common PPTX patterns
SLIDE_LAYOUT_REL_CONSTRAINT = RelationshipExistConstraint(
    attribute="id",
    namespace=REL_NS,
    relationship_type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
)

SLIDE_MASTER_REL_CONSTRAINT = RelationshipExistConstraint(
    attribute="id",
    namespace=REL_NS,
    relationship_type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
)

THEME_REL_CONSTRAINT = RelationshipExistConstraint(
    attribute="id",
    namespace=REL_NS,
    relationship_type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
)
