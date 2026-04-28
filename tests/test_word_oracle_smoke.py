"""End-to-end smoke test for the Word roundtrip oracle.

Marked `requires_word_app` and skipped automatically when Microsoft Word
for Mac is not reachable. Run by hand on a configured developer machine.

Spec: `specs/011-word-roundtrip-oracle.md` Phase 1 exit gate.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _build_minimal_docx(tmp_path: Path, name: str, body_xml: str) -> Path:
    doc = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>{body_xml}</w:body>
</w:document>"""
    rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    doc_rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    p = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("_rels/.rels", rels)
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
    p.write_bytes(buf.getvalue())
    return p


@pytest.mark.requires_word_app
def test_clean_baseline_roundtrip_preserves_content(tmp_path: Path) -> None:
    """A canonical-order DOCX should roundtrip through Word with no repair
    dialog. This is the smoke test that proves the engine works at all."""
    from tools.oracle.diff import diff_property_fragment, extract_first
    from tools.oracle.word_roundtrip import roundtrip

    body = """\
<w:tbl>
  <w:tr>
    <w:trPr>
      <w:cantSplit/>
      <w:tblHeader/>
    </w:trPr>
    <w:tc><w:p/></w:tc>
  </w:tr>
</w:tbl>"""
    docx = _build_minimal_docx(tmp_path, "clean.docx", body)

    # Don't pass output_dir — pytest's tmp_path lives in /private/var/folders/
    # which is outside Word's sandbox. Let roundtrip() use its default
    # Documents-scoped staging directory.
    result = roundtrip(docx, timeout=60.0)

    assert result.output_path.exists()
    assert result.repair_dialog_seen is False, (
        "Canonical-order trPr unexpectedly triggered Word's repair dialog: "
        f"{result.repair_dialog_text!r}"
    )

    input_children = extract_first(str(result.input_path), "trPr", W)
    output_children = extract_first(str(result.output_path), "trPr", W)
    diff = diff_property_fragment("trPr", input_children, output_children)
    assert diff.verdict == "preserved", diff.summary


@pytest.mark.requires_word_app
def test_issue_3_repro_observed_outcome(tmp_path: Path) -> None:
    """Issue #3 exact repro (tblHeader before cantSplit) — observe what
    Word actually does with this input.

    Spec 010 Phase 1 ships a constraint that flags this ordering on the
    assumption (per issue #3) that Word triggers its repair dialog. The
    oracle's job is to verify or refute that assumption.

    Observed in the first run of this test (Word for Mac M365 16.89.1,
    AppleScript open + close-with-save path): NO repair dialog, trPr
    children preserved as-is. Possible explanations:
      - The AppleScript open path is more permissive than UI File>Open
      - Synthetic minimal DOCX bypasses validation that richer content triggers
      - Word version drift between issue #3's reporter and this machine

    This test does not assert the verdict — it documents the observation
    and would surface a regression if Word's behavior on this exact
    input changed. Spec 010 Phase 4 needs to investigate the gap before
    the constraint can be considered oracle-grade.
    """
    from tools.oracle.diff import diff_property_fragment, extract_first
    from tools.oracle.word_roundtrip import roundtrip

    body = """\
<w:tbl>
  <w:tr>
    <w:trPr>
      <w:tblHeader/>
      <w:cantSplit/>
    </w:trPr>
    <w:tc><w:p/></w:tc>
  </w:tr>
</w:tbl>"""
    docx = _build_minimal_docx(tmp_path, "issue3.docx", body)
    result = roundtrip(docx, timeout=60.0)

    input_children = extract_first(str(result.input_path), "trPr", W)
    output_children = extract_first(str(result.output_path), "trPr", W)
    diff = diff_property_fragment("trPr", input_children, output_children)

    # Print the observation so it's visible in pytest output and recorded
    # in CI logs (when run on a developer machine).
    print(
        f"\nWord {result.word_version} on issue #3 repro: "
        f"repair_dialog_seen={result.repair_dialog_seen}, "
        f"verdict={diff.verdict}, "
        f"summary={diff.summary}"
    )
    # Only assert the engine completed — both files exist and we got a verdict.
    assert result.output_path.exists()
    assert diff.verdict in ("preserved", "repaired", "missing")
