"""Semantic reference constraints for Open XML validation.

These constraints validate ID references, index references, and
uniqueness requirements within the XML document.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.semantic.attributes import SemanticConstraint

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext


@dataclass
class IndexReferenceConstraint(SemanticConstraint):
    """Validates that an index attribute references a valid index.

    Used for cases like referencing slide numbers, color scheme indices, etc.
    """

    attribute: str
    max_index_attribute: str | None = None  # Attribute containing max valid index
    max_index_xpath: str | None = None  # XPath to count elements for max index
    zero_based: bool = True
    namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        try:
            index = int(element.attrib[attr_name])
        except ValueError:
            context.add_semantic_error(
                f"Index attribute '{self.attribute}' must be an integer",
                node=self.attribute,
            )
            return False

        min_index = 0 if self.zero_based else 1

        if index < min_index:
            context.add_semantic_error(
                f"Index '{self.attribute}' value {index} is less than minimum {min_index}",
                node=self.attribute,
            )
            return False

        # Check max index if specified
        max_index = None

        if self.max_index_attribute:
            max_attr_name = (
                f"{{{self.namespace}}}{self.max_index_attribute}"
                if self.namespace
                else self.max_index_attribute
            )
            if max_attr_name in element.attrib:
                try:
                    max_index = int(element.attrib[max_attr_name])
                except ValueError:
                    pass

        if self.max_index_xpath and max_index is None:
            # Count elements using XPath
            try:
                result = element.xpath(self.max_index_xpath)
                if isinstance(result, list):
                    max_index = len(result) - (1 if self.zero_based else 0)
            except Exception:
                pass

        if max_index is not None and index > max_index:
            context.add_semantic_error(
                f"Index '{self.attribute}' value {index} exceeds maximum {max_index}",
                node=self.attribute,
            )
            return False

        return True


@dataclass
class ReferenceExistConstraint(SemanticConstraint):
    """Validates that an ID reference points to an existing element.

    Used for cases like referencing shape IDs, placeholder IDs, etc.
    """

    attribute: str
    target_xpath: str  # XPath to find target elements with matching ID
    id_attribute: str = "id"  # Attribute name on target elements
    namespace: str | None = None
    target_namespace: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        ref_id = element.attrib[attr_name]

        # Get the root element to search from
        root = element.getroottree().getroot()

        # Build the XPath to find matching elements
        target_id_name = (
            f"{{{self.target_namespace}}}{self.id_attribute}"
            if self.target_namespace
            else self.id_attribute
        )

        # Search for element with matching ID
        try:
            xpath = f"{self.target_xpath}[@{self.id_attribute}='{ref_id}']"
            matches = root.xpath(xpath, namespaces={"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})
            if not matches:
                context.add_semantic_error(
                    f"Reference '{self.attribute}' value '{ref_id}' does not match any element",
                    node=self.attribute,
                )
                return False
        except Exception:
            # XPath evaluation failed, skip validation
            pass

        return True


@dataclass
class UniqueAttributeValueConstraint(SemanticConstraint):
    """Validates that attribute values are unique within a scope.

    Used for ensuring IDs are unique, names don't conflict, etc.
    """

    attribute: str
    scope_xpath: str = "."  # XPath defining scope for uniqueness
    namespace: str | None = None
    case_sensitive: bool = True
    element_tag: str | None = None

    def validate(
        self, element: etree._Element, context: ValidationContext
    ) -> bool:
        attr_name = f"{{{self.namespace}}}{self.attribute}" if self.namespace else self.attribute

        if attr_name not in element.attrib:
            return True

        value = element.attrib[attr_name]
        if not self.case_sensitive:
            value = value.lower()

        # Find scope root
        try:
            if self.scope_xpath == ".":
                scope_root = element.getparent()
            else:
                results = element.xpath(self.scope_xpath)
                if not results:
                    return True
                scope_root = results[0] if isinstance(results, list) else results
        except Exception:
            return True

        # Count elements with same attribute value
        count = 0
        if self.element_tag:
            root = element.getroottree().getroot()
            candidates = root.iter(self.element_tag)
        else:
            if scope_root is None:
                return True
            candidates = scope_root.iter()

        for elem in candidates:
            if attr_name in elem.attrib:
                elem_value = elem.attrib[attr_name]
                if not self.case_sensitive:
                    elem_value = elem_value.lower()
                if elem_value == value:
                    count += 1

        if count > 1:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' value '{value}' is not unique "
                f"(found {count} occurrences)",
                node=self.attribute,
            )
            return False

        return True


class IdTracker:
    """Tracks IDs across a document for uniqueness validation."""

    def __init__(self) -> None:
        self._ids: dict[str, set[str]] = {}  # scope -> set of IDs

    def add_id(self, scope: str, id_value: str) -> bool:
        """Add an ID and return True if it was unique.

        Args:
            scope: The scope (e.g., part URI) for uniqueness checking.
            id_value: The ID value.

        Returns:
            True if the ID was unique in this scope, False if duplicate.
        """
        if scope not in self._ids:
            self._ids[scope] = set()

        if id_value in self._ids[scope]:
            return False

        self._ids[scope].add(id_value)
        return True

    def has_id(self, scope: str, id_value: str) -> bool:
        """Check if an ID exists in a scope."""
        return scope in self._ids and id_value in self._ids[scope]

    def clear(self, scope: str | None = None) -> None:
        """Clear tracked IDs.

        Args:
            scope: If provided, only clear IDs in this scope.
                   Otherwise, clear all IDs.
        """
        if scope is None:
            self._ids.clear()
        elif scope in self._ids:
            del self._ids[scope]


def validate_unique_ids(
    element: etree._Element,
    id_attribute: str,
    context: "ValidationContext",
    tracker: IdTracker,
    scope: str | None = None,
) -> bool:
    """Validate that all ID attributes in an element tree are unique.

    Args:
        element: Root element to validate.
        id_attribute: Name of the ID attribute to check.
        context: Validation context.
        tracker: ID tracker for uniqueness checking.
        scope: Scope key for uniqueness (defaults to part URI).

    Returns:
        True if all IDs are unique.
    """
    if scope is None:
        scope = context.part_uri

    valid = True

    for elem in element.iter():
        if id_attribute in elem.attrib:
            id_value = elem.attrib[id_attribute]
            if not tracker.add_id(scope, id_value):
                context.add_semantic_error(
                    f"Duplicate ID '{id_value}'",
                    node=id_attribute,
                )
                valid = False

    return valid
