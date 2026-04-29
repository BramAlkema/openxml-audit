"""Tests for the Excel canonical-form validator (Spec 029, 0.7.3).

`ExcelCanonicalFormValidator` detects patterns Excel will rewrite on
save — silent canonicalization that the v0.7.2 baseline showed
happens on every TokenMoulds-emitted workbook.

This test suite covers Excel_InlineStrCells (the 0.7.3 ship); future
checks (chart externalLink materialization, attribute-order
canonicalization) get their own tests when shipped.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from openxml_audit import FileFormat, OpenXmlValidator
from openxml_audit.errors import SourceClass


REPO_ROOT = Path(__file__).resolve().parents[1]


def _build_xlsx(
    target: Path,
    *,
    sheet_xml: bytes,
    shared_strings_xml: bytes | None = None,
    extra_parts: dict[str, bytes] | None = None,
) -> Path:
    """Build a minimal but valid `.xlsx` for testing.

    `sheet_xml` is the body of `xl/worksheets/sheet1.xml`. When
    `shared_strings_xml` is None, no `xl/sharedStrings.xml` part is
    emitted (the trigger condition for Excel_InlineStrCells).
    """
    target.parent.mkdir(parents=True, exist_ok=True)
    parts: dict[str, bytes] = {
        "[Content_Types].xml": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b'<Override PartName="/xl/workbook.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            b'<Override PartName="/xl/worksheets/sheet1.xml" '
            b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            + (
                b'<Override PartName="/xl/sharedStrings.xml" '
                b'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                if shared_strings_xml is not None
                else b""
            )
            + b'</Types>'
        ),
        "_rels/.rels": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            b'Target="xl/workbook.xml"/>'
            b'</Relationships>'
        ),
        "xl/workbook.xml": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            b'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
            b'</workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            b'Target="worksheets/sheet1.xml"/>'
            + (
                b'<Relationship Id="rId2" '
                b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                b'Target="sharedStrings.xml"/>'
                if shared_strings_xml is not None
                else b""
            )
            + b'</Relationships>'
        ),
        "xl/worksheets/sheet1.xml": sheet_xml,
    }
    if shared_strings_xml is not None:
        parts["xl/sharedStrings.xml"] = shared_strings_xml
    if extra_parts:
        parts.update(extra_parts)

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return target


_SHEET_WITH_INLINE_STRINGS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<sheetData>'
    b'<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>'
    b'<row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c></row>'
    b'<row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c></row>'
    b'</sheetData>'
    b'</worksheet>'
)

_SHEET_WITH_SHARED_STRING_REFS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<sheetData>'
    b'<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
    b'<row r="2"><c r="A2" t="s"><v>1</v></c></row>'
    b'</sheetData>'
    b'</worksheet>'
)

_SHEET_NO_STRING_CELLS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<sheetData>'
    b'<row r="1"><c r="A1"><v>42</v></c></row>'
    b'<row r="2"><c r="A2"><v>3.14</v></c></row>'
    b'</sheetData>'
    b'</worksheet>'
)

_POPULATED_SHARED_STRINGS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'count="2" uniqueCount="2">'
    b'<si><t>Header</t></si>'
    b'<si><t>Q1</t></si>'
    b'</sst>'
)

_EMPTY_SHARED_STRINGS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'count="0" uniqueCount="0"/>'
)


def _validate(path: Path) -> list:
    v = OpenXmlValidator(file_format=FileFormat.OFFICE_2019, strict=False)
    return v.validate(path).errors


def _inline_str_findings(errors: list) -> list:
    return [e for e in errors if e.id == "Excel_InlineStrCells"]


# -- positive cases — the check fires ----------------------------------------


def test_inlinestr_cells_without_shared_strings_part_flagged(tmp_path: Path) -> None:
    """The dominant v0.7.2 corpus condition: inlineStr cells AND no
    `xl/sharedStrings.xml` at all. Should flag the worksheet."""
    pkg = _build_xlsx(
        tmp_path / "no_sst.xlsx",
        sheet_xml=_SHEET_WITH_INLINE_STRINGS,
        shared_strings_xml=None,
    )
    findings = _inline_str_findings(_validate(pkg))
    assert len(findings) == 1
    f = findings[0]
    assert f.source_class is SourceClass.EXCEL_APP_COMPAT
    assert f.severity.value == "warning"
    assert f.part_uri == "/xl/worksheets/sheet1.xml"
    assert "3 inline-string cell" in f.description


def test_inlinestr_cells_with_empty_shared_strings_part_flagged(tmp_path: Path) -> None:
    """Empty SST is also the trigger: Excel won't migrate INTO an
    empty table, it grows one. Same canonical-form gap."""
    pkg = _build_xlsx(
        tmp_path / "empty_sst.xlsx",
        sheet_xml=_SHEET_WITH_INLINE_STRINGS,
        shared_strings_xml=_EMPTY_SHARED_STRINGS,
    )
    findings = _inline_str_findings(_validate(pkg))
    assert len(findings) == 1


# -- negative cases — the check does NOT fire --------------------------------


def test_no_inlinestr_no_finding(tmp_path: Path) -> None:
    """Sheet has no inline-string cells at all — nothing to flag."""
    pkg = _build_xlsx(
        tmp_path / "no_strings.xlsx",
        sheet_xml=_SHEET_NO_STRING_CELLS,
    )
    findings = _inline_str_findings(_validate(pkg))
    assert findings == []


def test_inlinestr_with_populated_shared_strings_no_finding(tmp_path: Path) -> None:
    """Edge case: sheet has BOTH inlineStr cells AND a populated SST.
    Excel still rewrites the inlineStr cells, but per current scope
    we only flag the no-SST case (the dominant TokenMoulds pattern)
    to keep the WARNING signal-to-noise high. This test locks that
    decision."""
    pkg = _build_xlsx(
        tmp_path / "both.xlsx",
        sheet_xml=_SHEET_WITH_INLINE_STRINGS,
        shared_strings_xml=_POPULATED_SHARED_STRINGS,
    )
    findings = _inline_str_findings(_validate(pkg))
    assert findings == []


def test_shared_string_refs_only_no_finding(tmp_path: Path) -> None:
    """Sheet uses `<c t="s">` cells (canonical) — no flag."""
    pkg = _build_xlsx(
        tmp_path / "canonical.xlsx",
        sheet_xml=_SHEET_WITH_SHARED_STRING_REFS,
        shared_strings_xml=_POPULATED_SHARED_STRINGS,
    )
    findings = _inline_str_findings(_validate(pkg))
    assert findings == []


# -- corpus smoke — the v0.7.2 reference fixture ----------------------------


def test_v072_tokenmoulds_acme_us_xlsx_flagged() -> None:
    """The `data/corpus/tokenmoulds_v0.7.2/excel/acme-us.xlsx` is
    the canonical positive fixture — the file the v0.7.2 oracle
    baseline showed Excel rewrites every save. The validator should
    flag it now, before any Excel save, with the count matching
    what the file actually contains (9 inlineStr cells per the
    v0.7.2 manual probe)."""
    fixture = REPO_ROOT / "data/corpus/tokenmoulds_v0.7.2/excel/acme-us.xlsx"
    if not fixture.exists():
        pytest.skip("v0.7.2 corpus not committed yet")
    findings = _inline_str_findings(_validate(fixture))
    assert len(findings) == 1
    assert "9 inline-string cell" in findings[0].description
    assert findings[0].source_class is SourceClass.EXCEL_APP_COMPAT


def test_globex_xlsx_also_flagged() -> None:
    """Second brand variant in the v0.7.2 corpus — same emitter
    output shape, same finding."""
    fixture = REPO_ROOT / "data/corpus/tokenmoulds_v0.7.2/excel/globex-gb.xlsx"
    if not fixture.exists():
        pytest.skip("v0.7.2 corpus not committed yet")
    findings = _inline_str_findings(_validate(fixture))
    assert len(findings) == 1
