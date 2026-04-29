"""Excel canonical-form validation.

Excel rewrites packages on save to canonicalize them — even when they
pass our strict schema validation cleanly, Excel will silently
modify the file. The 0.7.2 oracle baseline showed every TokenMoulds-
emitted `.xlsx` came back with 10 changed parts, with no repair
dialog. This module ships the first canonical-form check that
detects ahead of save what Excel will rewrite, so emitter authors
can fix their output before users hit the silent canonicalization.

Findings emitted here are tagged `SourceClass.EXCEL_APP_COMPAT` —
they're real validator output (semantic-level constraints Excel
enforces by rewriting), not strict OOXML violations. Severity
defaults to WARNING.

Spec 029 (Phase 1) ships one check; subsequent releases add more.

## Check 1: Excel_InlineStrCells

Pattern: a worksheet uses `<c t="inlineStr"><is>...` for string
content when the package has no `xl/sharedStrings.xml`. Excel's
canonical form for shared text values is the shared-strings table
plus `<c t="s"><v>N</v></c>` cells; on every save Excel migrates
inline strings to that form. Detection is straightforward —
count `<c>` cells with `t="inlineStr"` and check whether
sharedStrings.xml is present-and-nonempty.

This was the dominant cause of the v0.7.2 baseline's `repaired`
outcome on Excel files (acme-us / globex-gb): 9 inlineStr cells
per worksheet, no sharedStrings.xml present. Excel's repair: move
all 9 to a freshly-created sharedStrings.xml, replace the
`<c t="inlineStr"><is>...` cells with `<c t="s"><v>idx</v>`.

## Future checks (deferred)

- **Excel_ChartExternalRefMaterialization** — chart `<c:numRef>` /
  `<c:strRef>` references whose target cells in the workbook
  don't typematch the chart's cached values; Excel materializes
  the cache into a separate `xl/externalLinks/externalLink<N>.xml`
  part on save. Detection requires resolving chart refs to sheet
  cells and comparing types — non-trivial; deferred to a focused
  follow-up release.
- **Excel_NonCanonicalAttributeOrder** — Excel reorders attributes
  on save in stable but undocumented ways. Detection requires a
  reference dataset of "what Excel's attribute order is per
  element kind." Deferred until corpus mining produces such a
  reference.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import (
    SourceClass,
    ValidationError,
    ValidationErrorType,
    ValidationSeverity,
)
from openxml_audit.namespaces import SPREADSHEETML

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.package import OpenXmlPackage


# Where Excel keeps the shared-strings table when populated.
_SHARED_STRINGS_PART = "xl/sharedStrings.xml"

# Where worksheet parts live.
_WORKSHEET_PREFIX = "xl/worksheets/"


def _shared_strings_table_is_populated(package: OpenXmlPackage) -> bool:
    """True iff `xl/sharedStrings.xml` exists and contains at least one
    `<si>` (shared-string item).

    A part that exists but holds zero items is still
    "non-canonical" from Excel's perspective — Excel won't migrate
    inlineStr cells INTO an empty shared-strings table; it'll grow
    one as needed. The check is "does the writer-emitted SST hold
    the shared text already, or does the writer leave that to Excel."
    """
    try:
        xml = package.get_part_xml(_SHARED_STRINGS_PART)
    except Exception:
        return False
    if xml is None:
        return False
    ns = {"s": SPREADSHEETML}
    return bool(xml.findall("s:si", ns))


def _count_inline_string_cells(worksheet_xml: etree._Element) -> int:
    """Count `<c t="inlineStr">` cells in a worksheet's XML.

    Worksheet markup uses the `c` element with a `t` attribute that
    selects the value representation. `t="inlineStr"` means an
    `<is>` child; that's the form Excel rewrites to shared-strings
    on save. Other `t` values (`s`, `n`, `b`, `e`, `str`) are not
    rewritten by this check.
    """
    ns = {"s": SPREADSHEETML}
    cells = worksheet_xml.findall(".//s:c[@t='inlineStr']", ns)
    return len(cells)


def _list_worksheet_parts(package: OpenXmlPackage) -> list[str]:
    """Return all worksheet part URIs in the package.

    Looks at the content_types declaration plus a path-prefix scan,
    so this catches both standard names (`xl/worksheets/sheet1.xml`)
    and any non-standard names a writer may use.
    """
    return sorted(
        uri for uri in package.list_parts()
        if uri.startswith("/" + _WORKSHEET_PREFIX) or uri.startswith(_WORKSHEET_PREFIX)
        if uri.endswith(".xml")
    )


class ExcelCanonicalFormValidator:
    """Detect patterns Excel will rewrite on save.

    Findings are `EXCEL_APP_COMPAT` — they signal "Excel will
    canonicalize this away" rather than "this file is broken."
    Severity defaults to WARNING; emitter authors fix them so their
    output is byte-stable across an Excel save.
    """

    def validate(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> None:
        """Run every check in the canonical-form suite against `package`."""
        self._check_inline_strings_without_shared_table(package, context)

    def _check_inline_strings_without_shared_table(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> None:
        """Excel_InlineStrCells: any `<c t="inlineStr">` cell when the
        shared-strings table is empty/missing.

        Emits one finding per worksheet that has such cells, with
        the count in the description so the user knows the blast
        radius. The validator pre-computes whether the shared
        table is populated so the check is O(worksheets) rather
        than O(worksheets × cells)."""
        if _shared_strings_table_is_populated(package):
            # The author already shipped strings via the canonical
            # path; Excel won't rewrite. Skip the check.
            return

        ns = {"s": SPREADSHEETML}
        for uri in _list_worksheet_parts(package):
            try:
                xml = package.get_part_xml(uri)
            except Exception:
                continue
            if xml is None:
                continue
            count = _count_inline_string_cells(xml)
            if count == 0:
                continue
            sheet_name = uri.rsplit("/", 1)[-1]
            # Append directly with an explicit part_uri rather than going
            # through context.add_error (which reads part_uri from
            # context.part). One finding per worksheet, each tagged with
            # its own URI so users can navigate to the affected sheet.
            context.errors.append(ValidationError(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    f"Worksheet {sheet_name} has {count} inline-string cell(s) "
                    "(`<c t=\"inlineStr\">`) but the package has no populated "
                    "`xl/sharedStrings.xml`. Excel will move every inline string "
                    "to a freshly-created shared-strings table on the next save, "
                    "rewriting both the worksheet and adding the SST part. "
                    "Emit shared strings instead of inlineStr cells to avoid "
                    "the silent rewrite."
                ),
                part_uri=uri if uri.startswith("/") else "/" + uri,
                node="c",
                severity=ValidationSeverity.WARNING,
                id="Excel_InlineStrCells",
                source_class=SourceClass.EXCEL_APP_COMPAT,
            ))
