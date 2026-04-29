"""Tests for the Excel roundtrip oracle (Spec 021, 0.6.5)."""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.xlsx_repair_oracle import (  # noqa: E402
    PartFingerprint,
    RoundtripObservation,
    _FINGERPRINTED_PREFIXES,
    _fingerprint_xlsx,
    _is_fingerprinted_part,
    _stage_root,
    _to_jsonable,
)
from openxml_audit.xlsx import osa as xlsx_osa  # noqa: E402


def _excel_available() -> bool:
    return sys.platform == "darwin" and xlsx_osa.EXCEL_APP_BUNDLE.exists()


REQUIRES_EXCEL = pytest.mark.skipif(
    not _excel_available(),
    reason="Microsoft Excel not installed (or non-darwin); "
           "Excel oracle tests skipped",
)


def _build_synthetic_xlsx(target: Path) -> Path:
    """Minimum-viable .xlsx for harness-only tests."""
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        zf.writestr(
            "_rels/.rels",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        zf.writestr(
            "xl/workbook.xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        )
        zf.writestr(
            "xl/worksheets/sheet1.xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        )
        zf.writestr("xl/media/image1.png", b"\x89PNG fake")  # excluded
    return target


# -- always-on tests ---------------------------------------------------------


def test_fingerprinted_prefixes_includes_known_canonical_parts() -> None:
    assert "xl/workbook.xml" in _FINGERPRINTED_PREFIXES
    assert "xl/worksheets/" in _FINGERPRINTED_PREFIXES
    assert "[Content_Types].xml" in _FINGERPRINTED_PREFIXES


def test_is_fingerprinted_part_canonical_parts() -> None:
    assert _is_fingerprinted_part("xl/workbook.xml")
    assert _is_fingerprinted_part("xl/worksheets/sheet1.xml")
    assert _is_fingerprinted_part("xl/styles.xml")
    assert _is_fingerprinted_part("[Content_Types].xml")


def test_is_fingerprinted_part_excludes_media() -> None:
    assert not _is_fingerprinted_part("xl/media/image1.png")
    assert not _is_fingerprinted_part("docProps/thumbnail.jpeg")
    assert not _is_fingerprinted_part("xl/printerSettings/printerSettings1.bin")


def test_fingerprint_xlsx_skips_excluded_parts(tmp_path: Path) -> None:
    xlsx = _build_synthetic_xlsx(tmp_path / "synthetic.xlsx")
    fps = _fingerprint_xlsx(xlsx)
    names = {fp.name for fp in fps}
    assert "xl/workbook.xml" in names
    assert "xl/worksheets/sheet1.xml" in names
    assert "[Content_Types].xml" in names
    assert "xl/media/image1.png" not in names
    for fp in fps:
        assert isinstance(fp, PartFingerprint)
        assert len(fp.sha256) == 64


def test_fingerprint_xlsx_returns_empty_for_missing_file(tmp_path: Path) -> None:
    assert _fingerprint_xlsx(tmp_path / "nope.xlsx") == []


def test_to_jsonable_summary_counts_outcomes_and_dialogs() -> None:
    obs = [
        RoundtripObservation(source_relpath="a.xlsx", outcome="preserved",
                             duration_seconds=1.0),
        RoundtripObservation(
            source_relpath="b.xlsx", outcome="repaired", duration_seconds=2.0,
            repair_dialog_seen=True, repair_dialog_text="found a problem",
            changed_parts=["xl/worksheets/sheet1.xml"],
        ),
    ]
    payload = _to_jsonable(obs)
    assert payload["schema_version"] == 1
    assert payload["summary"] == {
        "total": 2,
        "preserved": 1,
        "repaired": 1,
        "crash": 0,
        "timeout": 0,
        "open_failed": 0,
        "missing_output": 0,
        "repair_dialog_seen": 1,
    }


def test_stage_root_honors_env_override(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("XLSX_ORACLE_STAGE", str(tmp_path / "custom"))
    root = _stage_root()
    assert root == tmp_path / "custom"
    assert root.is_dir()


def test_stage_root_default_is_documents_subdir(monkeypatch) -> None:
    monkeypatch.delenv("XLSX_ORACLE_STAGE", raising=False)
    root = _stage_root()
    assert root.parent.name == "Documents"
    assert root.name == ".xlsx_oracle_runs"


def test_xlsx_osa_exports_save_close_primitives() -> None:
    expected = {
        "save_workbook",
        "close_workbook",
        "close_workbook_saving",
        "find_repair_dialog_text",
        "is_workbook_open",
        "list_open_workbook_names",
    }
    assert expected.issubset(set(xlsx_osa.__all__))
    for name in expected:
        assert callable(getattr(xlsx_osa, name))


def test_xlsx_osa_repair_dialog_patterns_documented() -> None:
    patterns = xlsx_osa.REPAIR_DIALOG_PATTERNS
    assert len(patterns) >= 5
    assert "found a problem" in patterns
    assert "repair" in patterns


# -- Excel-required smoke ----------------------------------------------------


@REQUIRES_EXCEL
def test_excel_app_bundle_path_exists_when_install_detected() -> None:
    assert xlsx_osa.EXCEL_APP_BUNDLE.exists()
    assert xlsx_osa.EXCEL_APP_BUNDLE.is_dir()


@REQUIRES_EXCEL
def test_list_open_workbook_names_returns_list() -> None:
    """Sanity: AppleScript bridge to Excel's `workbooks` collection runs.

    Returns either an empty list (no workbooks open) or a list of names.
    Doesn't launch Excel just to enumerate."""
    result = xlsx_osa.list_open_workbook_names()
    assert isinstance(result, list)
    assert all(isinstance(n, str) for n in result)
