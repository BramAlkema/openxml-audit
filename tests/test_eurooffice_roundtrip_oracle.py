from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree

from openxml_audit.docx.oracle_starter_doc import build_minimal_docx
from openxml_audit.namespaces import WORDPROCESSINGML

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.eurooffice_roundtrip import (  # noqa: E402
    EditorPassEvidence,
    EditorTransportEvidence,
    _to_jsonable,
    observe,
)

W = f"{{{WORDPROCESSINGML}}}"


def _body(text: str) -> etree._Element:
    body = etree.Element(f"{W}body", nsmap={"w": WORDPROCESSINGML})
    paragraph = etree.SubElement(body, f"{W}p")
    run = etree.SubElement(paragraph, f"{W}r")
    node = etree.SubElement(run, f"{W}t")
    node.text = text
    etree.SubElement(body, f"{W}sectPr")
    return body


def _evidence(output_path: Path) -> EditorTransportEvidence:
    digest = "a" * 64
    return EditorTransportEvidence(
        remote_filename="openxml-oracle-test.docx",
        browser_image="mcr.microsoft.com/playwright:v1.62.0-noble",
        uploaded_sha256=digest,
        inserted_sha256="b" * 64,
        final_sha256=digest,
        passes=(
            EditorPassEvidence(
                action="insert",
                opened=True,
                dirty=True,
                disconnect="browser_exit",
                page_url="http://127.0.0.1/example/editor",
                frame_url="http://127.0.0.1/web-apps/documenteditor",
            ),
            EditorPassEvidence(
                action="remove",
                opened=True,
                dirty=True,
                disconnect="browser_exit",
                page_url="http://127.0.0.1/example/editor",
                frame_url="http://127.0.0.1/web-apps/documenteditor",
            ),
        ),
    )


class _CopyClient:
    def roundtrip(self, input_path: Path, output_path: Path) -> EditorTransportEvidence:
        output_path.write_bytes(input_path.read_bytes())
        return _evidence(output_path)


class _DriftClient:
    def roundtrip(self, input_path: Path, output_path: Path) -> EditorTransportEvidence:
        build_minimal_docx(output_path, body=_body("Changed by editor"))
        return _evidence(output_path)


def test_observe_separates_editor_evidence_from_semantic_preservation(
    tmp_path: Path,
    monkeypatch,
) -> None:
    monkeypatch.setenv("EUROOFFICE_ORACLE_STAGE", str(tmp_path / "stage"))
    source = tmp_path / "source.docx"
    build_minimal_docx(source, body=_body("Unchanged"))

    observation = observe(source, client=_CopyClient())

    assert observation.outcome == "preserved"
    assert observation.semantic_comparison is not None
    assert observation.semantic_comparison["preserved"] is True
    assert observation.editor_evidence is not None
    assert len(observation.editor_evidence["passes"]) == 2


def test_observe_reports_semantic_drift_even_when_editor_transport_succeeds(
    tmp_path: Path,
    monkeypatch,
) -> None:
    monkeypatch.setenv("EUROOFFICE_ORACLE_STAGE", str(tmp_path / "stage"))
    source = tmp_path / "source.docx"
    build_minimal_docx(source, body=_body("Original"))

    observation = observe(source, client=_DriftClient())

    assert observation.outcome == "semantic_drift"
    assert observation.semantic_comparison is not None
    assert "content_blocks" in observation.semantic_comparison["changed_features"]
    assert observation.editor_evidence is not None


def test_report_rolls_up_changed_features(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("EUROOFFICE_ORACLE_STAGE", str(tmp_path / "stage"))
    source = tmp_path / "source.docx"
    build_minimal_docx(source, body=_body("Original"))
    observation = observe(source, client=_DriftClient())

    report = _to_jsonable([observation])

    assert report["engine"] == "eurooffice-editor"
    assert report["summary"]["semantic_drift"] == 1
    assert report["summary"]["changed_feature_counts"]["content_blocks"] == 1
