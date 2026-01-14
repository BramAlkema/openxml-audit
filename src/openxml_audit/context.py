"""Validation context for tracking state during validation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import (
    FileFormat,
    ValidationError,
    ValidationErrorType,
    ValidationSeverity,
)

if TYPE_CHECKING:
    from openxml_audit.package import OpenXmlPackage
    from openxml_audit.parts import OpenXmlPart


@dataclass
class ElementInfo:
    """Information about an element being validated."""

    element: etree._Element
    path: str  # XPath-like path to element
    depth: int


class ValidationStack:
    """Stack for tracking element traversal during validation."""

    def __init__(self) -> None:
        self._stack: list[ElementInfo] = []

    def push(self, element: etree._Element, name: str) -> None:
        """Push an element onto the stack."""
        depth = len(self._stack)
        if self._stack:
            parent_path = self._stack[-1].path
            path = f"{parent_path}/{name}"
        else:
            path = f"/{name}"

        self._stack.append(ElementInfo(element=element, path=path, depth=depth))

    def pop(self) -> ElementInfo | None:
        """Pop an element from the stack."""
        if self._stack:
            return self._stack.pop()
        return None

    @property
    def current(self) -> ElementInfo | None:
        """Get the current element."""
        return self._stack[-1] if self._stack else None

    @property
    def current_path(self) -> str:
        """Get the current element path."""
        return self._stack[-1].path if self._stack else ""

    @property
    def depth(self) -> int:
        """Get the current depth."""
        return len(self._stack)

    def __len__(self) -> int:
        return len(self._stack)


@dataclass
class ValidationContext:
    """Context for validation operations.

    Tracks the current state during validation including:
    - Current package and part being validated
    - Element traversal stack
    - Collected errors
    - Configuration settings
    """

    package: OpenXmlPackage | None = None
    part: OpenXmlPart | None = None
    file_format: FileFormat = FileFormat.OFFICE_2019
    max_errors: int = 1000
    strict: bool = True
    errors: list[ValidationError] = field(default_factory=list)
    _stack: ValidationStack = field(default_factory=ValidationStack)

    @property
    def part_uri(self) -> str:
        """Get the current part URI."""
        return self.part.uri if self.part else ""

    @property
    def current_path(self) -> str:
        """Get the current element path."""
        return self._stack.current_path

    @property
    def current_element(self) -> etree._Element | None:
        """Get the current element being validated."""
        info = self._stack.current
        return info.element if info else None

    def push_element(self, element: etree._Element) -> None:
        """Push an element onto the traversal stack."""
        # Get local name without namespace
        tag = element.tag
        if tag.startswith("{"):
            tag = tag.split("}")[-1]
        self._stack.push(element, tag)

    def pop_element(self) -> None:
        """Pop an element from the traversal stack."""
        self._stack.pop()

    def add_error(
        self,
        error_type: ValidationErrorType,
        description: str,
        node: str | None = None,
        related_node: str | None = None,
        severity: ValidationSeverity = ValidationSeverity.ERROR,
        error_id: str = "",
    ) -> None:
        """Add a validation error using current context."""
        if self.max_errors > 0 and self.error_count >= self.max_errors:
            return

        if not self.strict and severity == ValidationSeverity.ERROR:
            if error_type != ValidationErrorType.PACKAGE:
                severity = ValidationSeverity.WARNING

        error = ValidationError(
            error_type=error_type,
            description=description,
            part_uri=self.part_uri,
            path=self.current_path,
            node=node,
            related_node=related_node,
            severity=severity,
            id=error_id,
        )
        self.errors.append(error)

    def add_schema_error(
        self,
        description: str,
        node: str | None = None,
        error_id: str = "",
    ) -> None:
        """Add a schema validation error."""
        self.add_error(
            error_type=ValidationErrorType.SCHEMA,
            description=description,
            node=node,
            error_id=error_id,
        )

    def add_semantic_error(
        self,
        description: str,
        node: str | None = None,
        error_id: str = "",
    ) -> None:
        """Add a semantic validation error."""
        self.add_error(
            error_type=ValidationErrorType.SEMANTIC,
            description=description,
            node=node,
            error_id=error_id,
        )

    @property
    def error_count(self) -> int:
        """Get the number of errors collected."""
        return sum(1 for e in self.errors if e.severity == ValidationSeverity.ERROR)

    @property
    def should_stop(self) -> bool:
        """Check if we should stop collecting errors."""
        if self.max_errors == 0:
            return False
        return self.error_count >= self.max_errors

    def set_part(self, part: OpenXmlPart) -> None:
        """Set the current part being validated."""
        self.part = part
        # Clear the element stack when changing parts
        self._stack = ValidationStack()

    def clear_errors(self) -> None:
        """Clear all collected errors."""
        self.errors.clear()


class ElementContext:
    """Context manager for element traversal."""

    def __init__(self, context: ValidationContext, element: etree._Element):
        self._context = context
        self._element = element

    def __enter__(self) -> ElementContext:
        self._context.push_element(self._element)
        return self

    def __exit__(self, exc_type: object, exc_val: object, exc_tb: object) -> None:
        self._context.pop_element()
