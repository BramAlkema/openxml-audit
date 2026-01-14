"""Word document structure validation."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import ValidationError
from openxml_audit.namespaces import REL_FOOTER, REL_HEADER, WORDPROCESSINGML

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import DocumentPart


class DocumentValidator:
    """Validate Word document structure and relationships."""

    def __init__(self) -> None:
        self._ns = {"w": WORDPROCESSINGML}
        self._rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def validate(
        self, part: "DocumentPart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a Word document part."""
        context.set_part(part)
        errors: list[ValidationError] = []

        xml = part.xml
        if xml is None:
            return errors

        with ElementContext(context, xml):
            self._validate_root(xml, context)
            self._validate_body(xml, context)
            self._validate_header_footer_refs(xml, part, context)

        errors.extend(context.errors)
        return errors

    def _validate_root(self, xml: etree._Element, context: "ValidationContext") -> None:
        expected_tag = f"{{{WORDPROCESSINGML}}}document"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'document', got '{xml.tag}'",
            )

    def _validate_body(self, xml: etree._Element, context: "ValidationContext") -> None:
        body = xml.find("w:body", self._ns)
        if body is None:
            context.add_schema_error(
                "Document is missing w:body element",
                node="body",
            )

    def _validate_header_footer_refs(
        self,
        xml: etree._Element,
        part: "DocumentPart",
        context: "ValidationContext",
    ) -> None:
        for tag, rel_type in (
            ("headerReference", REL_HEADER),
            ("footerReference", REL_FOOTER),
        ):
            for ref in xml.findall(f".//w:{tag}", self._ns):
                rel_id = ref.get(f"{{{self._rel_ns}}}id", "")
                if not rel_id:
                    context.add_schema_error(
                        f"{tag} is missing r:id attribute",
                        node=tag,
                    )
                    continue

                rel = part.relationships.get_by_id(rel_id)
                if rel is None:
                    context.add_semantic_error(
                        f"{tag} relationship '{rel_id}' not found",
                        node="r:id",
                    )
                elif rel.type != rel_type:
                    context.add_semantic_error(
                        f"{tag} relationship '{rel_id}' should be {rel_type}, got '{rel.type}'",
                        node="r:id",
                    )
