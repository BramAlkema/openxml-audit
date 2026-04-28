"""Tests for the Word compatibility ordering check (spec 010, Phase 1)."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from openxml_audit import OpenXmlValidator
from openxml_audit.errors import ValidationSeverity
from openxml_audit.word.compat import (
    CONSTRAINT_TABLE,
    ChildSequence,
    find_first_out_of_order,
)

# --- Subsequence engine tests ---------------------------------------------


CANONICAL = ("a", "b", "c", "d", "e")


def test_empty_observed_passes() -> None:
    assert find_first_out_of_order([], CANONICAL) is None


def test_in_order_passes() -> None:
    assert find_first_out_of_order(["a", "b", "c"], CANONICAL) is None


def test_in_order_with_skips_passes() -> None:
    assert find_first_out_of_order(["a", "c", "e"], CANONICAL) is None


def test_single_reorder_fails() -> None:
    result = find_first_out_of_order(["b", "a"], CANONICAL)
    assert result == (1, 0)  # 'a' at idx 1 should appear before 'b' at idx 0


def test_late_reorder_reports_first_violation() -> None:
    # 'a','c' is fine, then 'b' breaks
    result = find_first_out_of_order(["a", "c", "b"], CANONICAL)
    assert result == (2, 1)


def test_repeated_child_in_order_passes() -> None:
    # Repeats are allowed if non-decreasing in canonical position
    assert find_first_out_of_order(["a", "a", "b", "b"], CANONICAL) is None


def test_unknown_child_silently_skipped() -> None:
    # Children not in canonical are not our concern
    assert find_first_out_of_order(["a", "x", "b"], CANONICAL) is None


def test_unknown_does_not_mask_real_violation() -> None:
    result = find_first_out_of_order(["b", "x", "a"], CANONICAL)
    assert result == (2, 0)  # 'a' at idx 2 should appear before 'b' at idx 0


# --- Constraint-table integrity --------------------------------------------


def test_constraint_table_well_formed() -> None:
    assert CONSTRAINT_TABLE, "constraint table must not be empty"
    for parent_tag, entry in CONSTRAINT_TABLE.items():
        assert isinstance(entry, ChildSequence)
        assert entry.parent_tag == parent_tag, (
            f"entry parent_tag {entry.parent_tag!r} doesn't match key {parent_tag!r}"
        )
        assert parent_tag.startswith("{") and "}" in parent_tag, (
            f"parent_tag must use Clark notation: {parent_tag!r}"
        )
        assert entry.children, f"empty children tuple for {parent_tag}"
        assert len(set(entry.children)) == len(entry.children), (
            f"duplicate children in {parent_tag}: {entry.children}"
        )
        for child in entry.children:
            assert child.startswith("{") and "}" in child, (
                f"child must use Clark notation: {child!r}"
            )
        assert entry.spec_section, f"missing spec_section for {parent_tag}"
        assert entry.parent_local, f"missing parent_local for {parent_tag}"


def test_phase2_constraints_present() -> None:
    """Phase 2 covers tblPr, tcPr, sectPr — all corpus-validated."""
    from openxml_audit.namespaces import WORDPROCESSINGML

    for prop in ("tblPr", "tcPr", "sectPr"):
        parent = f"{{{WORDPROCESSINGML}}}{prop}"
        assert parent in CONSTRAINT_TABLE, f"missing {prop} entry"
        assert CONSTRAINT_TABLE[parent].children


def test_ct_trpr_canonical_order_includes_known_children() -> None:
    """Sanity check: the CT_TrPr entry covers the children the issue lists."""
    from openxml_audit.namespaces import WORDPROCESSINGML

    parent = f"{{{WORDPROCESSINGML}}}trPr"
    entry = CONSTRAINT_TABLE[parent]
    locals_in_order = [
        c.split("}", 1)[1] for c in entry.children if c.startswith(f"{{{WORDPROCESSINGML}}}")
    ]
    # cantSplit must precede tblHeader (the issue's repro)
    assert locals_in_order.index("cantSplit") < locals_in_order.index("tblHeader")
    # trPrChange comes near the end, after track-change ins/del
    assert locals_in_order.index("ins") < locals_in_order.index("del")
    assert locals_in_order.index("del") < locals_in_order.index("trPrChange")


# --- Integration tests against built DOCX inputs --------------------------


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

_CT = """\
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

_DOC_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""


def _build_docx(tmp_path: Path, name: str, body_xml: str) -> Path:
    doc = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>{body_xml}</w:body>
</w:document>"""
    files = {
        "_rels/.rels": _RELS,
        "[Content_Types].xml": _CT,
        "word/document.xml": doc,
        "word/_rels/document.xml.rels": _DOC_RELS,
    }
    path = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p, content in files.items():
            zf.writestr(p, content)
    path.write_bytes(buf.getvalue())
    return path


def _trpr(children_xml: str) -> str:
    """Wrap the given trPr children in a one-row table."""
    return f"""\
<w:tbl>
  <w:tr>
    <w:trPr>{children_xml}</w:trPr>
  </w:tr>
</w:tbl>"""


def test_in_order_trpr_produces_no_warning(tmp_path: Path) -> None:
    """cantSplit before tblHeader matches ECMA-376 §17.4.79 — no warning."""
    body = _trpr("<w:cantSplit/><w:tblHeader/>")
    pkg = _build_docx(tmp_path, "in_order.docx", body)

    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [
        e for e in result.errors if "trPr child" in e.description
    ]
    assert matches == [], [e.description for e in matches]


def test_issue_3_repro_fires_warning(tmp_path: Path) -> None:
    """Issue #3 exact repro: tblHeader before cantSplit triggers a WARNING."""
    body = _trpr("<w:tblHeader/><w:cantSplit/>")
    pkg = _build_docx(tmp_path, "out_of_order.docx", body)

    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [
        e for e in result.errors
        if "trPr child" in e.description and "cantSplit" in e.description
    ]
    assert matches, [e.description for e in result.errors]
    assert all(e.severity == ValidationSeverity.WARNING for e in matches), (
        [(e.severity, e.description) for e in matches]
    )
    assert any("§17.4.79" in e.description for e in matches), (
        [e.description for e in matches]
    )
    assert any("tblHeader" in e.description for e in matches)


def test_tblpr_reorder_fires_warning(tmp_path: Path) -> None:
    """Phase 2: tblPr — placing tblBorders before tblW (which precedes it
    canonically) should fire WARNING."""
    body = """\
<w:tbl>
  <w:tblPr>
    <w:tblBorders/>
    <w:tblW w:w="0" w:type="auto"/>
  </w:tblPr>
  <w:tr><w:tc><w:p/></w:tc></w:tr>
</w:tbl>"""
    pkg = _build_docx(tmp_path, "tblpr_reorder.docx", body)
    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [e for e in result.errors if "tblPr child" in e.description]
    assert matches, [e.description for e in result.errors]
    assert all(e.severity == ValidationSeverity.WARNING for e in matches)
    assert any("§17.4.60" in e.description for e in matches)


def test_tcpr_reorder_fires_warning(tmp_path: Path) -> None:
    """Phase 2: tcPr — shd before tcBorders is a reordering."""
    body = """\
<w:tbl>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:shd w:val="clear"/>
        <w:tcBorders/>
      </w:tcPr>
      <w:p/>
    </w:tc>
  </w:tr>
</w:tbl>"""
    pkg = _build_docx(tmp_path, "tcpr_reorder.docx", body)
    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [e for e in result.errors if "tcPr child" in e.description]
    assert matches, [e.description for e in result.errors]
    assert any("tcBorders" in e.description for e in matches)
    assert any("§17.4.70" in e.description for e in matches)


def test_sectpr_reorder_fires_warning(tmp_path: Path) -> None:
    """Phase 2: sectPr — pgMar before pgSz is a reordering."""
    body = """\
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:pPr>
</w:p>"""
    pkg = _build_docx(tmp_path, "sectpr_reorder.docx", body)
    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [e for e in result.errors if "sectPr child" in e.description]
    assert matches, [e.description for e in result.errors]
    assert any("pgSz" in e.description for e in matches)


def test_phase2_in_order_inputs_pass(tmp_path: Path) -> None:
    """In-order tblPr/tcPr/sectPr children produce no Phase 2 warnings."""
    body = """\
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/>
    </w:sectPr>
  </w:pPr>
</w:p>
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders/>
  </w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcBorders/>
        <w:shd w:val="clear"/>
      </w:tcPr>
      <w:p/>
    </w:tc>
  </w:tr>
</w:tbl>"""
    pkg = _build_docx(tmp_path, "phase2_clean.docx", body)
    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [
        e for e in result.errors
        if any(prop in e.description for prop in ("tblPr child", "tcPr child", "sectPr child"))
    ]
    assert matches == [], [e.description for e in matches]


def test_unknown_child_in_trpr_is_not_flagged(tmp_path: Path) -> None:
    """Children not in CT_TrPr canonical (e.g., extension elements) are
    silently skipped — ordering checks are scoped to known children."""
    body = _trpr(
        '<w:cantSplit/>'
        '<w:extLst><w:ext xmlns:custom="urn:x" custom:foo="1"/></w:extLst>'
        '<w:tblHeader/>'
    )
    pkg = _build_docx(tmp_path, "unknown.docx", body)

    result = OpenXmlValidator(schema_validation=False).validate(pkg)
    matches = [e for e in result.errors if "trPr child" in e.description]
    assert matches == [], [e.description for e in matches]


@pytest.mark.skipif(
    pytest.importorskip("docx", reason="python-docx not installed") is None,
    reason="python-docx not available",
)
def test_python_docx_default_produces_no_ordering_warnings(tmp_path: Path) -> None:
    """Regression control: a default python-docx document with a table must
    not trigger any Phase 1 ordering warnings (false-positive guard)."""
    from docx import Document

    p = tmp_path / "python_docx.docx"
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    # Touch row properties so trPr definitely appears
    for row in table.rows:
        row.height_rule = 1  # exact height — adds trHeight to trPr
    doc.save(str(p))

    result = OpenXmlValidator(schema_validation=False).validate(p)
    matches = [e for e in result.errors if "trPr child" in e.description]
    assert matches == [], [e.description for e in matches]
