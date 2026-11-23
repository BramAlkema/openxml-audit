"""Build minimal Word documents for calibration probes.

Pure lxml + `openxml_audit.builder.PackageBuilder` — no python-docx.
Caller supplies a `<w:body>` element to exercise a single feature (field
code, content control, track-change, etc.); this module provides only
the minimum wrapping required for Word to open the file.
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from openxml_audit.builder import PackageBuilder
from openxml_audit.namespaces import REL_OFFICE_DOCUMENT, WORDPROCESSINGML

__all__ = ["MINIMAL_BODY_XML", "build_minimal_docx"]


_W = f"{{{WORDPROCESSINGML}}}"

_CT_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
_CT_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"
_CT_XML = "application/xml"


MINIMAL_BODY_XML = b"""<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p/>
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
             w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
    <w:docGrid w:linePitch="360"/>
  </w:sectPr>
</w:body>
"""


def build_minimal_docx(
    output_path: Path | str,
    *,
    body: etree._Element | None = None,
) -> None:
    """Write a minimal `.docx` to `output_path`.

    If `body` is provided, it is used as the `<w:body>` element; otherwise
    a minimal empty-paragraph body with a standard US-Letter sectPr is used.
    """
    builder = PackageBuilder()
    builder.add_default_type("rels", _CT_RELATIONSHIPS)
    builder.add_default_type("xml", _CT_XML)
    builder.add_relationship("/", "rId1", REL_OFFICE_DOCUMENT, "word/document.xml")
    builder.add_part(
        "/word/document.xml",
        _build_document_xml(body),
        content_type=_CT_DOCUMENT,
    )
    builder.write(output_path)


def _build_document_xml(body: etree._Element | None) -> bytes:
    nsmap = {"w": WORDPROCESSINGML}
    document = etree.Element(f"{_W}document", nsmap=nsmap)
    if body is None:
        body = etree.fromstring(MINIMAL_BODY_XML)
    document.append(body)
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
