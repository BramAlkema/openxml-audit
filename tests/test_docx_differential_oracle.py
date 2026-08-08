from __future__ import annotations

from pathlib import Path

from lxml import etree

from openxml_audit.docx.differential_oracle import build_docx_differential_report
from openxml_audit.docx.oracle_starter_doc import build_minimal_docx
from openxml_audit.namespaces import WORDPROCESSINGML

W = f"{{{WORDPROCESSINGML}}}"


def _build(path: Path, text: str) -> None:
    body = etree.Element(f"{W}body", nsmap={"w": WORDPROCESSINGML})
    paragraph = etree.SubElement(body, f"{W}p")
    run = etree.SubElement(paragraph, f"{W}r")
    node = etree.SubElement(run, f"{W}t")
    node.text = text
    etree.SubElement(body, f"{W}sectPr")
    build_minimal_docx(path, body=body)


def test_differential_matrix_reports_portable_and_divergent_features(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    eurooffice = tmp_path / "eurooffice.docx"
    google = tmp_path / "google.docx"
    _build(source, "Original")
    _build(eurooffice, "Original")
    _build(google, "Changed")

    report = build_docx_differential_report(
        source,
        {"eurooffice": eurooffice, "google_docs": google},
    )

    assert report["all_targets_semantically_preserved"] is False
    assert "security_surface" in report["portable_features"]
    assert "content_blocks" in report["divergent_features"]
    content = next(
        entry for entry in report["feature_matrix"] if entry["feature"] == "content_blocks"
    )
    assert content["targets"]["eurooffice"]["preserved"] is True
    assert content["targets"]["google_docs"]["preserved"] is False


def test_differential_requires_two_targets(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    _build(source, "Original")

    try:
        build_docx_differential_report(source, {"only": source})
    except ValueError as exc:
        assert "at least two" in str(exc)
    else:
        raise AssertionError("expected a two-target guard")
