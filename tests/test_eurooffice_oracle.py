"""Tests for the Euro-Office conversion oracle (Spec 038)."""

from __future__ import annotations

import sys
from pathlib import Path

from openxml_audit.eurooffice import ConversionResult

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.eurooffice_conversion_oracle import (  # noqa: E402
    build_report,
    observe,
)


class _FakeClient:
    def __init__(self, output: bytes) -> None:
        self.output = output
        self.calls: list[dict[str, str]] = []

    def healthcheck(self) -> bool:
        return True

    def version(self) -> str:
        return "9.3.3"

    def convert(self, **kwargs: str) -> ConversionResult:
        self.calls.append(kwargs)
        return ConversionResult("https://result.example.test/converted")

    def download(self, file_url: str) -> bytes:
        return self.output


def test_native_same_format_conversion_is_diffed(
    minimal_pptx: Path,
    tmp_path: Path,
) -> None:
    client = _FakeClient(minimal_pptx.read_bytes())
    observation = observe(
        minimal_pptx,
        client=client,
        server_version="9.3.3",
        source_base_url="https://files.example.test/corpus/",
        work_root=tmp_path / "work",
    )

    assert observation.outcome == "preserved"
    assert observation.format_mode == "native-edit"
    assert observation.source_valid is True
    assert observation.source_error_count == 0
    assert observation.target_valid is True
    assert observation.changed_parts == []
    assert observation.sha256_in == observation.sha256_out
    assert client.calls[0]["source_url"].endswith("/minimal.pptx")
    assert client.calls[0]["target_format"] == "pptx"


def test_odf_edit_path_is_reported_as_lossy_conversion(
    minimal_odt: Path,
    tmp_path: Path,
) -> None:
    valid_docx = REPO_ROOT / "data" / "corpus" / "tokenmoulds_v0.7.2" / "word" / "acme-us.docx"
    client = _FakeClient(valid_docx.read_bytes())
    observation = observe(
        minimal_odt,
        client=client,
        server_version="9.3.3",
        source_base_url="https://files.example.test/corpus/",
        work_root=tmp_path / "work",
    )

    assert observation.outcome == "converted"
    assert observation.format_mode == "lossy-edit"
    assert observation.source_format == "odt"
    assert observation.target_format == "docx"
    assert observation.target_valid is True
    assert any("converts odt to docx" in note for note in observation.notes)


def test_odg_is_conversion_only(minimal_pptx: Path, tmp_path: Path) -> None:
    source = tmp_path / "drawing.odg"
    source.write_bytes(b"synthetic ODG source")
    observation = observe(
        source,
        client=_FakeClient(minimal_pptx.read_bytes()),
        source_base_url="https://files.example.test/",
        work_root=tmp_path / "work",
    )
    assert observation.outcome == "converted"
    assert observation.format_mode == "view-only"
    assert observation.target_format == "pptx"
    assert any("not an editable" in note for note in observation.notes)


def test_invalid_same_format_output_still_records_diff(
    minimal_pptx: Path,
    invalid_pptx_missing_presentation: Path,
    tmp_path: Path,
) -> None:
    observation = observe(
        minimal_pptx,
        client=_FakeClient(invalid_pptx_missing_presentation.read_bytes()),
        source_base_url="https://files.example.test/",
        work_root=tmp_path / "work",
    )
    assert observation.outcome == "invalid_output"
    assert observation.source_valid is True
    assert observation.target_valid is False
    assert "ppt/presentation.xml" in observation.removed_parts


def test_unsupported_format_does_not_call_server(tmp_path: Path) -> None:
    source = tmp_path / "database.odb"
    source.write_bytes(b"not sent")
    client = _FakeClient(b"")
    observation = observe(source, client=client)
    assert observation.outcome == "unsupported"
    assert observation.format_mode == "unsupported"
    assert client.calls == []


def test_report_keeps_conversion_evidence_scope(minimal_pptx: Path, tmp_path: Path) -> None:
    observation = observe(
        minimal_pptx,
        client=_FakeClient(minimal_pptx.read_bytes()),
        source_base_url="https://files.example.test/",
        work_root=tmp_path / "work",
    )
    report = build_report([observation], server_version="9.3.3")
    assert report["engine"] == "eurooffice"
    assert report["evidence_scope"] == "Document Server conversion endpoint; not browser editing"
    assert report["upstream"]["document_server_release"] == "9.3.3"
    assert report["summary"] == {"total": 1, "outcomes": {"preserved": 1}}


def test_dispatcher_has_eurooffice_aliases() -> None:
    from openxml_audit.oracle.__main__ import _DISPATCH

    assert _DISPATCH["eurooffice"] is _DISPATCH["euro-office"]
    assert _DISPATCH["eurooffice"] is _DISPATCH["euro"]
