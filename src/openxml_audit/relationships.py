"""Relationship handling for OPC packages."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.namespaces import RELATIONSHIPS

if TYPE_CHECKING:
    from collections.abc import Iterator


@dataclass
class Relationship:
    """An OPC relationship between parts."""

    id: str  # Relationship ID (e.g., "rId1")
    type: str  # Relationship type URI
    target: str  # Target path (relative or absolute)
    target_mode: str = "Internal"  # "Internal" or "External"

    @property
    def is_external(self) -> bool:
        """Check if this is an external relationship."""
        return self.target_mode == "External"

    def resolve_target(self, source_uri: str) -> str:
        """Resolve the target path relative to the source part.

        Args:
            source_uri: The URI of the part containing this relationship.
                       For package relationships, use "/".

        Returns:
            Absolute path to the target part within the package.
        """
        if self.is_external:
            return self.target

        if self.target.startswith("/"):
            # Already absolute
            return self.target

        # Get the directory of the source part
        source_path = PurePosixPath(source_uri)
        source_dir = source_path.parent

        # Resolve relative path
        target_path = source_dir / self.target
        # Normalize (resolve .. and .)
        parts = []
        for part in target_path.parts:
            if part == "..":
                if parts and parts[-1] != "/":
                    parts.pop()
            elif part != ".":
                parts.append(part)

        return "/" + "/".join(p for p in parts if p != "/")


class RelationshipCollection:
    """Collection of relationships for a part or package."""

    def __init__(self, source_uri: str = "/"):
        self._relationships: dict[str, Relationship] = {}
        self._source_uri = source_uri

    def add(self, rel: Relationship) -> None:
        """Add a relationship to the collection."""
        self._relationships[rel.id] = rel

    def get_by_id(self, rel_id: str) -> Relationship | None:
        """Get a relationship by its ID."""
        return self._relationships.get(rel_id)

    def get_by_type(self, rel_type: str) -> Iterator[Relationship]:
        """Get all relationships of a specific type."""
        for rel in self._relationships.values():
            if rel.type == rel_type:
                yield rel

    def get_first_by_type(self, rel_type: str) -> Relationship | None:
        """Get the first relationship of a specific type."""
        return next(self.get_by_type(rel_type), None)

    def resolve_target(self, rel_id: str) -> str | None:
        """Resolve the target path for a relationship ID."""
        rel = self.get_by_id(rel_id)
        if rel is None:
            return None
        return rel.resolve_target(self._source_uri)

    def __iter__(self) -> Iterator[Relationship]:
        return iter(self._relationships.values())

    def __len__(self) -> int:
        return len(self._relationships)

    def __contains__(self, rel_id: str) -> bool:
        return rel_id in self._relationships

    @classmethod
    def from_xml(cls, xml_content: bytes, source_uri: str = "/") -> RelationshipCollection:
        """Parse relationships from XML content.

        Args:
            xml_content: The raw XML bytes of the .rels file.
            source_uri: The URI of the part these relationships belong to.

        Returns:
            A RelationshipCollection containing all parsed relationships.
        """
        collection = cls(source_uri)

        try:
            root = etree.fromstring(xml_content)
        except etree.XMLSyntaxError:
            return collection

        # Handle namespace
        ns = {"r": RELATIONSHIPS}

        for rel_elem in root.findall("r:Relationship", ns):
            rel_id = rel_elem.get("Id", "")
            rel_type = rel_elem.get("Type", "")
            target = rel_elem.get("Target", "")
            target_mode = rel_elem.get("TargetMode", "Internal")

            if rel_id and rel_type:
                rel = Relationship(
                    id=rel_id,
                    type=rel_type,
                    target=target,
                    target_mode=target_mode,
                )
                collection.add(rel)

        return collection


def get_rels_path(part_uri: str) -> str:
    """Get the path to the .rels file for a given part.

    Args:
        part_uri: The URI of the part (e.g., "/ppt/presentation.xml").

    Returns:
        The path to the corresponding .rels file.
        For "/" returns "/_rels/.rels".
        For "/ppt/presentation.xml" returns "/ppt/_rels/presentation.xml.rels".
    """
    if part_uri == "/":
        return "/_rels/.rels"

    path = PurePosixPath(part_uri)
    rels_dir = path.parent / "_rels"
    rels_name = path.name + ".rels"

    return str(rels_dir / rels_name)
