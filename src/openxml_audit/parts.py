"""OpenXml Part handling."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.relationships import RelationshipCollection

if TYPE_CHECKING:
    from openxml_audit.package import OpenXmlPackage


class OpenXmlPart:
    """Represents a part within an OPC package."""

    def __init__(self, package: OpenXmlPackage, uri: str, content_type: str | None = None):
        self._package = package
        self._uri = uri
        self._content_type = content_type
        self._xml: etree._Element | None = None
        self._relationships: RelationshipCollection | None = None
        self._loaded = False

    @property
    def uri(self) -> str:
        """Get the URI of this part within the package."""
        return self._uri

    @property
    def content_type(self) -> str | None:
        """Get the content type of this part."""
        if self._content_type is None:
            self._content_type = self._package.content_types.get_content_type(self._uri)
        return self._content_type

    @property
    def xml(self) -> etree._Element | None:
        """Get the parsed XML content of this part.

        Returns None if the part doesn't exist or isn't valid XML.
        """
        if not self._loaded:
            self._xml = self._package.get_part_xml(self._uri)
            self._loaded = True
        return self._xml

    @property
    def raw_content(self) -> bytes | None:
        """Get the raw bytes content of this part."""
        return self._package.get_part_content(self._uri)

    @property
    def relationships(self) -> RelationshipCollection:
        """Get the relationships for this part."""
        if self._relationships is None:
            self._relationships = self._package.get_part_relationships(self._uri)
        return self._relationships

    @property
    def exists(self) -> bool:
        """Check if this part exists in the package."""
        return self._package.has_part(self._uri)

    def get_related_part(self, rel_id: str) -> OpenXmlPart | None:
        """Get a part related to this one by relationship ID.

        Args:
            rel_id: The relationship ID (e.g., "rId1").

        Returns:
            The related OpenXmlPart, or None if not found.
        """
        target = self.relationships.resolve_target(rel_id)
        if target is None:
            return None
        return OpenXmlPart(self._package, target)

    def get_related_parts_by_type(self, rel_type: str) -> list[OpenXmlPart]:
        """Get all parts related to this one by relationship type.

        Args:
            rel_type: The relationship type URI.

        Returns:
            List of related OpenXmlParts.
        """
        parts = []
        for rel in self.relationships.get_by_type(rel_type):
            target = rel.resolve_target(self._uri)
            parts.append(OpenXmlPart(self._package, target))
        return parts


class PresentationPart(OpenXmlPart):
    """The main presentation part (ppt/presentation.xml)."""

    CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    )

    def __init__(self, package: OpenXmlPackage, uri: str = "/ppt/presentation.xml"):
        super().__init__(package, uri, self.CONTENT_TYPE)

    @property
    def slide_ids(self) -> list[tuple[str, str]]:
        """Get list of (slide_id, rel_id) tuples for all slides."""
        from openxml_audit.namespaces import PRESENTATIONML

        slides = []
        xml = self.xml
        if xml is None:
            return slides

        ns = {"p": PRESENTATIONML}
        sld_id_lst = xml.find("p:sldIdLst", ns)
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall("p:sldId", ns):
                id_val = sld_id.get("id", "")
                # r:id attribute uses relationships namespace
                rel_id = sld_id.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                if id_val and rel_id:
                    slides.append((id_val, rel_id))

        return slides

    @property
    def slide_master_ids(self) -> list[tuple[str, str]]:
        """Get list of (id, rel_id) tuples for all slide masters."""
        from openxml_audit.namespaces import PRESENTATIONML

        masters = []
        xml = self.xml
        if xml is None:
            return masters

        ns = {"p": PRESENTATIONML}
        master_id_lst = xml.find("p:sldMasterIdLst", ns)
        if master_id_lst is not None:
            for master_id in master_id_lst.findall("p:sldMasterId", ns):
                id_val = master_id.get("id", "")
                rel_id = master_id.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                if rel_id:  # id is optional for masters
                    masters.append((id_val, rel_id))

        return masters


class SlidePart(OpenXmlPart):
    """A slide part (ppt/slides/slideN.xml)."""

    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    def __init__(self, package: OpenXmlPackage, uri: str):
        super().__init__(package, uri, self.CONTENT_TYPE)


class SlideLayoutPart(OpenXmlPart):
    """A slide layout part (ppt/slideLayouts/slideLayoutN.xml)."""

    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"

    def __init__(self, package: OpenXmlPackage, uri: str):
        super().__init__(package, uri, self.CONTENT_TYPE)


class SlideMasterPart(OpenXmlPart):
    """A slide master part (ppt/slideMasters/slideMasterN.xml)."""

    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"

    def __init__(self, package: OpenXmlPackage, uri: str):
        super().__init__(package, uri, self.CONTENT_TYPE)


class ThemePart(OpenXmlPart):
    """A theme part (ppt/theme/themeN.xml)."""

    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml"

    def __init__(self, package: OpenXmlPackage, uri: str):
        super().__init__(package, uri, self.CONTENT_TYPE)


class DocumentPart(OpenXmlPart):
    """The main Word document part (word/document.xml)."""

    CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    )

    def __init__(self, package: OpenXmlPackage, uri: str = "/word/document.xml"):
        super().__init__(package, uri, self.CONTENT_TYPE)


class WorkbookPart(OpenXmlPart):
    """The main Excel workbook part (xl/workbook.xml)."""

    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

    def __init__(self, package: OpenXmlPackage, uri: str = "/xl/workbook.xml"):
        super().__init__(package, uri, self.CONTENT_TYPE)

    @property
    def sheet_ids(self) -> list[tuple[str, str, str]]:
        """Get list of (sheet_id, rel_id, name) tuples for all sheets."""
        from openxml_audit.namespaces import SPREADSHEETML

        sheets: list[tuple[str, str, str]] = []
        xml = self.xml
        if xml is None:
            return sheets

        ns = {"s": SPREADSHEETML}
        sheet_list = xml.find("s:sheets", ns)
        if sheet_list is None:
            return sheets

        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        for sheet in sheet_list.findall("s:sheet", ns):
            sheet_id = sheet.get("sheetId", "")
            rel_id = sheet.get(f"{{{rel_ns}}}id", "")
            name = sheet.get("name", "")
            sheets.append((sheet_id, rel_id, name))

        return sheets
