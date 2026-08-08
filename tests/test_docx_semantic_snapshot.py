from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from openxml_audit.docx.oracle_starter_doc import build_minimal_docx
from openxml_audit.docx.semantic_snapshot import compare_docx_semantics, snapshot_docx
from openxml_audit.namespaces import WORDPROCESSINGML

W = f"{{{WORDPROCESSINGML}}}"


def _body(text: str) -> etree._Element:
    body = etree.Element(f"{W}body", nsmap={"w": WORDPROCESSINGML})
    paragraph = etree.SubElement(body, f"{W}p")
    run = etree.SubElement(paragraph, f"{W}r")
    node = etree.SubElement(run, f"{W}t")
    node.text = text
    section = etree.SubElement(body, f"{W}sectPr")
    page_size = etree.SubElement(section, f"{W}pgSz")
    page_size.set(f"{W}w", "11906")
    page_size.set(f"{W}h", "16838")
    return body


def test_identical_semantics_ignore_zip_metadata(tmp_path: Path) -> None:
    base = tmp_path / "base.docx"
    head = tmp_path / "head.docx"
    build_minimal_docx(base, body=_body("Portable content"))

    with zipfile.ZipFile(base) as source, zipfile.ZipFile(head, "w") as target:
        for info in reversed(source.infolist()):
            target.writestr(info.filename, source.read(info.filename))

    comparison = compare_docx_semantics(snapshot_docx(base), snapshot_docx(head))
    assert comparison.preserved is True
    assert comparison.changed_features == ()


def test_content_change_is_reported_without_collapsing_feature_matrix(tmp_path: Path) -> None:
    base = tmp_path / "base.docx"
    head = tmp_path / "head.docx"
    build_minimal_docx(base, body=_body("Before"))
    build_minimal_docx(head, body=_body("After"))

    comparison = compare_docx_semantics(snapshot_docx(base), snapshot_docx(head))
    assert comparison.preserved is False
    assert "content_blocks" in comparison.changed_features
    assert {feature.feature for feature in comparison.features} >= {
        "content_blocks",
        "style_semantics",
        "security_surface",
    }


def test_security_surface_tracks_external_relationships(tmp_path: Path) -> None:
    base = tmp_path / "base.docx"
    head = tmp_path / "head.docx"
    build_minimal_docx(base, body=_body("Same"))
    build_minimal_docx(head, body=_body("Same"))

    with zipfile.ZipFile(head, "a") as archive:
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.invalid/" TargetMode="External"/>
</Relationships>''',
        )

    comparison = compare_docx_semantics(snapshot_docx(base), snapshot_docx(head))
    assert "security_surface" in comparison.changed_features
