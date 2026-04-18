"""Minimal OPC (OOXML) package builder for calibration emitters.

Produces minimal valid Office packages where every element is chosen for
evidence-gathering, not product emission. Used by format-specific oracle
starters (`docx.oracle_starter_doc`, `xlsx.oracle_starter_book`, PPTX
oracle decks) to exercise single features for tier calibration.

ADR-002: calibration emitters are research infrastructure, not product
emission — they belong here, not in converter repos.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from openxml_audit.namespaces import CONTENT_TYPES, RELATIONSHIPS
from openxml_audit.relationships import get_rels_path

__all__ = ["PackageBuilder"]


@dataclass(frozen=True, slots=True)
class _Relationship:
    id: str
    type: str
    target: str


class PackageBuilder:
    """Assemble an OPC package from parts, content types, and relationships."""

    def __init__(self) -> None:
        self._parts: dict[str, bytes] = {}
        self._default_types: dict[str, str] = {}
        self._override_types: dict[str, str] = {}
        self._relationships: dict[str, list[_Relationship]] = {}

    def add_default_type(self, extension: str, content_type: str) -> None:
        """Register a default content type for a file extension."""
        self._default_types[extension] = content_type

    def add_part(
        self,
        part_uri: str,
        xml: bytes,
        *,
        content_type: str | None = None,
    ) -> None:
        """Add a part (XML file) to the package, optionally with an override content type."""
        part_uri = self._normalize_uri(part_uri)
        self._parts[part_uri] = xml
        if content_type is not None:
            self._override_types[part_uri] = content_type

    def add_relationship(
        self,
        source_uri: str,
        rel_id: str,
        rel_type: str,
        target: str,
    ) -> None:
        """Add a relationship from source_uri (use '/' for package-level rels) to target."""
        source_uri = self._normalize_uri(source_uri)
        self._relationships.setdefault(source_uri, []).append(
            _Relationship(rel_id, rel_type, target)
        )

    def to_bytes(self) -> bytes:
        """Assemble the package as a ZIP archive and return its bytes."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", self._build_content_types_xml())
            for source_uri, rels in self._relationships.items():
                rels_zip_path = get_rels_path(source_uri).lstrip("/")
                zf.writestr(rels_zip_path, _build_relationships_xml(rels))
            for part_uri, xml in self._parts.items():
                zf.writestr(part_uri.lstrip("/"), xml)
        return buffer.getvalue()

    def write(self, output_path: Path | str) -> None:
        """Write the package to disk."""
        Path(output_path).write_bytes(self.to_bytes())

    @staticmethod
    def _normalize_uri(uri: str) -> str:
        if uri == "/":
            return "/"
        if not uri.startswith("/"):
            return "/" + uri
        return uri

    def _build_content_types_xml(self) -> bytes:
        nsmap = {None: CONTENT_TYPES}
        root = etree.Element(f"{{{CONTENT_TYPES}}}Types", nsmap=nsmap)
        for extension, content_type in sorted(self._default_types.items()):
            default = etree.SubElement(root, f"{{{CONTENT_TYPES}}}Default")
            default.set("Extension", extension)
            default.set("ContentType", content_type)
        for part_uri, content_type in sorted(self._override_types.items()):
            override = etree.SubElement(root, f"{{{CONTENT_TYPES}}}Override")
            override.set("PartName", part_uri)
            override.set("ContentType", content_type)
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )


def _build_relationships_xml(rels: list[_Relationship]) -> bytes:
    nsmap = {None: RELATIONSHIPS}
    root = etree.Element(f"{{{RELATIONSHIPS}}}Relationships", nsmap=nsmap)
    for rel in rels:
        elem = etree.SubElement(root, f"{{{RELATIONSHIPS}}}Relationship")
        elem.set("Id", rel.id)
        elem.set("Type", rel.type)
        elem.set("Target", rel.target)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
