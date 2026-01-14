"""OPC (Open Packaging Conventions) package handling.

A PPTX file is an OPC package - a ZIP archive containing XML parts and relationships.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.core.package import ZipPackage
from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.namespaces import CONTENT_TYPES, REL_OFFICE_DOCUMENT
from openxml_audit.relationships import RelationshipCollection, get_rels_path

if TYPE_CHECKING:
    from collections.abc import Iterator


@dataclass
class ContentType:
    """A content type definition from [Content_Types].xml."""

    content_type: str
    part_name: str | None = None  # For Override elements
    extension: str | None = None  # For Default elements


@dataclass
class ContentTypes:
    """Content types from [Content_Types].xml."""

    defaults: dict[str, str] = field(default_factory=dict)  # extension -> content_type
    overrides: dict[str, str] = field(default_factory=dict)  # part_name -> content_type

    def get_content_type(self, part_name: str) -> str | None:
        """Get the content type for a part.

        First checks overrides, then falls back to extension-based defaults.
        """
        # Normalize part name (ensure leading slash)
        if not part_name.startswith("/"):
            part_name = "/" + part_name

        # Check overrides first
        if part_name in self.overrides:
            return self.overrides[part_name]

        # Fall back to extension-based default
        ext = Path(part_name).suffix.lstrip(".")
        return self.defaults.get(ext)

    @classmethod
    def from_xml(cls, xml_content: bytes) -> ContentTypes:
        """Parse [Content_Types].xml content."""
        ct = cls()

        try:
            root = etree.fromstring(xml_content)
        except etree.XMLSyntaxError:
            return ct

        ns = {"ct": CONTENT_TYPES}

        # Parse Default elements
        for default in root.findall("ct:Default", ns):
            ext = default.get("Extension", "")
            content_type = default.get("ContentType", "")
            if ext and content_type:
                ct.defaults[ext] = content_type

        # Parse Override elements
        for override in root.findall("ct:Override", ns):
            part_name = override.get("PartName", "")
            content_type = override.get("ContentType", "")
            if part_name and content_type:
                ct.overrides[part_name] = content_type

        return ct


class OpenXmlPackage(ZipPackage):
    """An Open XML package (PPTX, DOCX, XLSX).

    Provides access to the package structure, parts, relationships, and content types.
    """

    def __init__(self, path: str | Path):
        super().__init__(path)
        self._content_types: ContentTypes | None = None
        self._relationships: RelationshipCollection | None = None

    @property
    def content_types(self) -> ContentTypes:
        """Get the content types from [Content_Types].xml."""
        if self._content_types is None:
            self._content_types = self._load_content_types()
        return self._content_types

    @property
    def relationships(self) -> RelationshipCollection:
        """Get the package-level relationships from _rels/.rels."""
        if self._relationships is None:
            self._relationships = self._load_relationships("/")
        return self._relationships

    def _load_content_types(self) -> ContentTypes:
        """Load and parse [Content_Types].xml."""
        content = self.get_part_content("[Content_Types].xml")
        if content is None:
            self._errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description="Missing [Content_Types].xml",
                    severity=ValidationSeverity.ERROR,
                )
            )
            return ContentTypes()

        try:
            return ContentTypes.from_xml(content)
        except Exception as e:
            self._errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=f"Error parsing [Content_Types].xml: {e}",
                    severity=ValidationSeverity.ERROR,
                )
            )
            return ContentTypes()

    def _load_relationships(self, source_uri: str) -> RelationshipCollection:
        """Load relationships for a part or the package root."""
        rels_path = get_rels_path(source_uri)
        # Remove leading slash for ZIP lookup
        zip_path = rels_path.lstrip("/")

        content = self.get_part_content(zip_path)
        if content is None:
            if source_uri == "/":
                self._errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description="Missing _rels/.rels",
                        severity=ValidationSeverity.ERROR,
                    )
                )
            return RelationshipCollection(source_uri)

        return RelationshipCollection.from_xml(content, source_uri)

    def get_part_relationships(self, part_uri: str) -> RelationshipCollection:
        """Get the relationships for a specific part."""
        return self._load_relationships(part_uri)

    def list_parts(self) -> Iterator[str]:
        """List all parts in the package."""
        for name in super().list_parts():
            # Skip relationship files and content types
            if name.startswith("/_rels/") or "/_rels/" in name:
                continue
            if name == "/[Content_Types].xml":
                continue
            yield name

    def get_main_document_uri(self) -> str | None:
        """Get the URI of the main document part (presentation.xml for PPTX)."""
        rel = self.relationships.get_first_by_type(REL_OFFICE_DOCUMENT)
        if rel is None:
            return None
        return rel.resolve_target("/")

    def validate_structure(self) -> list[ValidationError]:
        """Perform basic structural validation of the package.

        Returns:
            List of validation errors found.
        """
        errors: list[ValidationError] = []

        # Check content types
        _ = self.content_types
        errors.extend(self._errors)

        # Check package relationships
        _ = self.relationships

        # Check for main document relationship
        main_doc = self.get_main_document_uri()
        if main_doc is None:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.RELATIONSHIP,
                    description="Missing main document relationship (officeDocument)",
                    part_uri="/_rels/.rels",
                    severity=ValidationSeverity.ERROR,
                )
            )
        elif not self.has_part(main_doc):
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=f"Main document part not found: {main_doc}",
                    part_uri=main_doc,
                    severity=ValidationSeverity.ERROR,
                )
            )

        return errors
