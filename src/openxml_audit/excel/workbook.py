"""Excel workbook structure validation."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import ValidationError
from openxml_audit.namespaces import (
    REL_CHARTSHEET,
    REL_DIALOGSHEET,
    REL_MACRO_SHEET,
    REL_WORKSHEET,
    SPREADSHEETML,
)

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import WorkbookPart


class WorkbookValidator:
    """Validate Excel workbook structure and sheet relationships."""

    def __init__(self) -> None:
        self._ns = {"s": SPREADSHEETML}
        self._rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        self._sheet_relationship_types = {
            REL_WORKSHEET,
            REL_CHARTSHEET,
            REL_DIALOGSHEET,
            REL_MACRO_SHEET,
        }

    def validate(
        self, part: "WorkbookPart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a workbook part."""
        context.set_part(part)
        errors: list[ValidationError] = []

        xml = part.xml
        if xml is None:
            return errors

        with ElementContext(context, xml):
            self._validate_root(xml, context)
            self._validate_sheet_list(xml, part, context)

        errors.extend(context.errors)
        return errors

    def _validate_root(self, xml: etree._Element, context: "ValidationContext") -> None:
        expected_tag = f"{{{SPREADSHEETML}}}workbook"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'workbook', got '{xml.tag}'",
            )

    def _validate_sheet_list(
        self,
        xml: etree._Element,
        part: "WorkbookPart",
        context: "ValidationContext",
    ) -> None:
        sheet_list = xml.find("s:sheets", self._ns)
        if sheet_list is None:
            context.add_schema_error(
                "Workbook is missing sheets element",
                node="sheets",
            )
            return

        sheets = sheet_list.findall("s:sheet", self._ns)
        if not sheets:
            context.add_schema_error(
                "Workbook contains no sheets",
                node="sheet",
            )
            return

        seen_ids: set[str] = set()
        seen_names: set[str] = set()

        for sheet in sheets:
            sheet_id = sheet.get("sheetId", "")
            name = sheet.get("name", "")

            if not sheet_id:
                context.add_schema_error(
                    "Sheet is missing sheetId attribute",
                    node="sheetId",
                )
            elif sheet_id in seen_ids:
                context.add_semantic_error(
                    f"Duplicate sheetId: {sheet_id}",
                    node="sheetId",
                )
            else:
                seen_ids.add(sheet_id)

            if not name:
                context.add_schema_error(
                    "Sheet is missing name attribute",
                    node="name",
                )
            elif name in seen_names:
                context.add_semantic_error(
                    f"Duplicate sheet name: {name}",
                    node="name",
                )
            else:
                seen_names.add(name)

            rel_id = sheet.get(f"{{{self._rel_ns}}}id", "")
            if not rel_id:
                context.add_schema_error(
                    "Sheet is missing r:id attribute",
                    node="r:id",
                )
                continue

            rel = part.relationships.get_by_id(rel_id)
            if rel is None:
                context.add_semantic_error(
                    f"Sheet relationship '{rel_id}' not found",
                    node="r:id",
                )
            elif rel.type not in self._sheet_relationship_types:
                context.add_semantic_error(
                    f"Sheet relationship '{rel_id}' has unexpected type '{rel.type}'",
                    node="r:id",
                )
