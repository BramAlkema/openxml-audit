"""Unit tests for the Word roundtrip oracle's pure-logic helpers.

Everything in this file runs without Word. Tests that depend on Word
itself live in `tests/test_word_oracle_smoke.py` and are gated by the
`requires_word_app` marker.

Spec: `specs/011-word-roundtrip-oracle.md`.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from tools.oracle.diff import (
    FragmentDiff,
    diff_property_fragment,
    extract_first,
    matches_repair_pattern,
    slugify_scenario_id,
)
from tools.oracle.word_window import REPAIR_DIALOG_PATTERNS

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _build_docx_with_trpr(tmp_path: Path, name: str, trpr_children_xml: str) -> Path:
    """Build a minimal valid DOCX with a single w:trPr containing the
    given children. Used to exercise extract_first against real packages."""
    doc = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:trPr>{trpr_children_xml}</w:trPr>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""
    rels = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    ct = """\
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    doc_rels = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""
    path = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("_rels/.rels", rels)
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
    path.write_bytes(buf.getvalue())
    return path


# --- diff_property_fragment ------------------------------------------------


def test_identical_children_preserved() -> None:
    result = diff_property_fragment("trPr", ["cantSplit", "tblHeader"], ["cantSplit", "tblHeader"])
    assert isinstance(result, FragmentDiff)
    assert result.verdict == "preserved"
    assert "identical" in result.summary


def test_reordered_children_repaired() -> None:
    result = diff_property_fragment("trPr", ["tblHeader", "cantSplit"], ["cantSplit", "tblHeader"])
    assert result.verdict == "repaired"
    assert "tblHeader" in result.summary
    assert "cantSplit" in result.summary


def test_input_missing_parent_reports_missing() -> None:
    result = diff_property_fragment("trPr", None, ["cantSplit"])
    assert result.verdict == "missing"
    assert "input" in result.summary


def test_output_missing_parent_reports_missing() -> None:
    result = diff_property_fragment("trPr", ["cantSplit"], None)
    assert result.verdict == "missing"
    assert "output" in result.summary


def test_added_child_repaired() -> None:
    result = diff_property_fragment(
        "trPr", ["cantSplit"], ["cantSplit", "tblHeader"]
    )
    assert result.verdict == "repaired"


def test_removed_child_repaired() -> None:
    result = diff_property_fragment(
        "trPr", ["cantSplit", "tblHeader"], ["cantSplit"]
    )
    assert result.verdict == "repaired"


# --- extract_first ---------------------------------------------------------


def test_extract_first_returns_children_in_order(tmp_path: Path) -> None:
    docx = _build_docx_with_trpr(
        tmp_path, "extract.docx", "<w:cantSplit/><w:tblHeader/>"
    )
    children = extract_first(str(docx), "trPr", W)
    assert children == ["cantSplit", "tblHeader"]


def test_extract_first_returns_none_when_parent_absent(tmp_path: Path) -> None:
    # Build a doc whose body has a paragraph but no table.
    doc = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body><w:p/></w:body>
</w:document>"""
    rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    p = tmp_path / "no_trpr.docx"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("_rels/.rels", rels)
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("word/document.xml", doc)
    p.write_bytes(buf.getvalue())
    assert extract_first(str(p), "trPr", W) is None


def test_extract_first_returns_first_when_multiple(tmp_path: Path) -> None:
    """When multiple parents exist, extract_first returns the first encountered."""
    doc = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:tbl>
      <w:tr><w:trPr><w:cantSplit/></w:trPr></w:tr>
      <w:tr><w:trPr><w:tblHeader/></w:trPr></w:tr>
    </w:tbl>
  </w:body>
</w:document>"""
    p = tmp_path / "multi.docx"
    buf = io.BytesIO()
    rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("_rels/.rels", rels)
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("word/document.xml", doc)
    p.write_bytes(buf.getvalue())
    children = extract_first(str(p), "trPr", W)
    assert children == ["cantSplit"]


# --- slugify_scenario_id ---------------------------------------------------


def test_slugify_scenario_id_basic() -> None:
    assert slugify_scenario_id("trPr", "baseline") == "trPr-baseline"


def test_slugify_scenario_id_with_parts() -> None:
    assert (
        slugify_scenario_id("trPr", "swap", "cantSplit", "tblHeader")
        == "trPr-swap-cantSplit-tblHeader"
    )


def test_slugify_scenario_id_strips_unsafe_chars() -> None:
    out = slugify_scenario_id("trPr", "swap", "child:foo", "child/bar")
    assert ":" not in out
    assert "/" not in out
    assert out == "trPr-swap-child_foo-child_bar"


def test_slugify_scenario_id_drops_empty_parts() -> None:
    assert slugify_scenario_id("trPr", "swap", "", "cantSplit") == "trPr-swap-cantSplit"


# --- matches_repair_pattern ------------------------------------------------


@pytest.mark.parametrize(
    "text",
    [
        "The document foo.docx contains unreadable content.",
        "Word found unreadable content in foo.docx. Do you want to recover the contents?",
        "Errors were detected in foo.docx",
    ],
)
def test_matches_repair_pattern_recognises_real_messages(text: str) -> None:
    assert matches_repair_pattern(text, REPAIR_DIALOG_PATTERNS)


def test_matches_repair_pattern_negative() -> None:
    assert not matches_repair_pattern("Document opened successfully.", REPAIR_DIALOG_PATTERNS)
    assert not matches_repair_pattern("", REPAIR_DIALOG_PATTERNS)


def test_matches_repair_pattern_case_insensitive() -> None:
    assert matches_repair_pattern("UNREADABLE CONTENT", REPAIR_DIALOG_PATTERNS)


# --- preflight (logic only) ------------------------------------------------


def test_preflight_module_importable() -> None:
    """Importing preflight must not raise even on non-macOS systems."""
    from tools.oracle import preflight

    assert hasattr(preflight, "check")


def test_preflight_returns_status_dataclass() -> None:
    from tools.oracle.preflight import PreflightStatus, check

    status = check()
    assert isinstance(status, PreflightStatus)
    assert isinstance(status.ok, bool)
    assert isinstance(status.issues, list)
