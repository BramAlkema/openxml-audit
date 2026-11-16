"""Build minimal Excel workbooks for calibration probes.

Pure lxml + `openxml_audit.builder.PackageBuilder` — no openpyxl.
Caller supplies a `<sheetData>` element to exercise a single feature
(dynamic array, shared-strings edge case, etc.); this module provides
only the minimum wrapping required for Excel to open the file.
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from openxml_audit.builder import PackageBuilder
from openxml_audit.namespaces import REL_OFFICE_DOCUMENT, SPREADSHEETML

__all__ = ["MINIMAL_SHEET_DATA_XML", "build_minimal_xlsx"]


_S = f"{{{SPREADSHEETML}}}"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
_CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_CT_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"
_CT_XML = "application/xml"

_REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


MINIMAL_SHEET_DATA_XML = b"""<sheetData xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"""


def build_minimal_xlsx(
    output_path: Path | str,
    *,
    sheet_data: etree._Element | None = None,
) -> None:
    """Write a minimal `.xlsx` to `output_path`.

    If `sheet_data` is provided, it is used as the `<sheetData>` element on
    sheet1; otherwise an empty `<sheetData/>` is written.
    """
    builder = PackageBuilder()
    builder.add_default_type("rels", _CT_RELATIONSHIPS)
    builder.add_default_type("xml", _CT_XML)
    builder.add_relationship("/", "rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml")
    builder.add_relationship("/xl/workbook.xml", "rId1", _REL_WORKSHEET, "worksheets/sheet1.xml")
    builder.add_part(
        "/xl/workbook.xml",
        _build_workbook_xml(),
        content_type=_CT_WORKBOOK,
    )
    builder.add_part(
        "/xl/worksheets/sheet1.xml",
        _build_worksheet_xml(sheet_data),
        content_type=_CT_WORKSHEET,
    )
    builder.write(output_path)


def _build_workbook_xml() -> bytes:
    nsmap = {None: SPREADSHEETML, "r": _R_NS}
    workbook = etree.Element(f"{_S}workbook", nsmap=nsmap)
    sheets = etree.SubElement(workbook, f"{_S}sheets")
    sheet = etree.SubElement(sheets, f"{_S}sheet")
    sheet.set("name", "Sheet1")
    sheet.set("sheetId", "1")
    sheet.set(f"{{{_R_NS}}}id", "rId1")
    return etree.tostring(workbook, xml_declaration=True, encoding="UTF-8", standalone=True)


def _build_worksheet_xml(sheet_data: etree._Element | None) -> bytes:
    nsmap = {None: SPREADSHEETML}
    worksheet = etree.Element(f"{_S}worksheet", nsmap=nsmap)
    if sheet_data is None:
        sheet_data = etree.fromstring(MINIMAL_SHEET_DATA_XML)
    worksheet.append(sheet_data)
    return etree.tostring(worksheet, xml_declaration=True, encoding="UTF-8", standalone=True)
