"""Tests for the PowerPoint roundtrip oracle (Spec 020, 0.6.4).

Split into:

  - **always-on**: pure-Python harness logic (no PowerPoint required).
  - **PowerPoint-required**: real-file roundtrip smoke. Skipped when
    Microsoft PowerPoint is not installed so CI without it still passes.

PowerPoint-required tests are also skipped on non-darwin platforms
(osascript / AppleScript only exist on macOS).
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.pptx_repair_oracle import (  # noqa: E402
    RoundtripObservation,
    _stage_root,
    _to_jsonable,
)
from openxml_audit.pptx import osa as pptx_osa  # noqa: E402


def _powerpoint_available() -> bool:
    return sys.platform == "darwin" and pptx_osa.POWERPOINT_APP_BUNDLE.exists()


REQUIRES_POWERPOINT = pytest.mark.skipif(
    not _powerpoint_available(),
    reason="Microsoft PowerPoint not installed (or non-darwin); "
           "PowerPoint oracle tests skipped",
)


def _build_synthetic_pptx(target: Path) -> Path:
    """Minimum-viable .pptx for harness-only tests (no real PowerPoint)."""
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        zf.writestr(
            "ppt/presentation.xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        )
        zf.writestr(
            "ppt/slides/slide1.xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        )
    return target


# -- always-on tests ---------------------------------------------------------


def test_to_jsonable_summary_counts_outcomes_and_dialogs() -> None:
    obs = [
        RoundtripObservation(source_relpath="a.pptx", outcome="preserved",
                             duration_seconds=1.0),
        RoundtripObservation(
            source_relpath="b.pptx", outcome="repaired", duration_seconds=2.0,
            repair_dialog_seen=True, repair_dialog_text="found a problem",
            changed_parts=["ppt/slides/slide1.xml"],
        ),
        RoundtripObservation(source_relpath="c.pptx", outcome="open_failed",
                             duration_seconds=0.5),
    ]
    payload = _to_jsonable(obs)
    assert payload["schema_version"] == 1
    assert payload["summary"] == {
        "total": 3,
        "preserved": 1,
        "repaired": 1,
        "crash": 0,
        "timeout": 0,
        "open_failed": 1,
        "missing_output": 0,
        "repair_dialog_seen": 1,
    }


def test_stage_root_creates_directory_under_documents(monkeypatch, tmp_path) -> None:
    """When PPTX_ORACLE_STAGE is set, _stage_root honors it."""
    monkeypatch.setenv("PPTX_ORACLE_STAGE", str(tmp_path / "custom"))
    root = _stage_root()
    assert root == tmp_path / "custom"
    assert root.is_dir()


def test_stage_root_default_is_documents_subdir(monkeypatch) -> None:
    """Default location is ~/Documents/.pptx_oracle_runs (App Sandbox-friendly)."""
    monkeypatch.delenv("PPTX_ORACLE_STAGE", raising=False)
    root = _stage_root()
    assert root.parent.name == "Documents"
    assert root.name == ".pptx_oracle_runs"


def test_pptx_osa_repair_dialog_patterns_documented() -> None:
    """REPAIR_DIALOG_PATTERNS exists, is non-empty, and lower-cased patterns
    cover the strings PowerPoint uses for its repair flow."""
    patterns = pptx_osa.REPAIR_DIALOG_PATTERNS
    assert len(patterns) >= 5
    assert "found a problem" in patterns
    assert "repair" in patterns


def test_pptx_osa_exports_save_close_primitives() -> None:
    """The osa layer must expose the primitives the oracle relies on.
    A regression in the `__all__` list would silently break importers."""
    expected = {
        "save_presentation",
        "close_presentation",
        "close_presentation_saving",
        "find_repair_dialog_text",
        "is_presentation_open",
        "list_open_presentation_names",
    }
    assert expected.issubset(set(pptx_osa.__all__))
    # And the names actually resolve to callables:
    for name in expected:
        assert callable(getattr(pptx_osa, name))


# -- PowerPoint-required smoke ----------------------------------------------


@REQUIRES_POWERPOINT
def test_powerpoint_app_bundle_path_exists_when_install_detected() -> None:
    assert pptx_osa.POWERPOINT_APP_BUNDLE.exists()
    assert pptx_osa.POWERPOINT_APP_BUNDLE.is_dir()


@REQUIRES_POWERPOINT
def test_list_open_presentation_names_returns_list() -> None:
    """Sanity: the AppleScript bridge to PowerPoint's `presentations`
    collection runs without raising. Returns either an empty list (no
    presentations open) or a list of names.

    We don't want to launch PowerPoint here just to enumerate; this
    test just verifies the wire works on whatever PowerPoint state
    happens to exist on the dev machine."""
    result = pptx_osa.list_open_presentation_names()
    assert isinstance(result, list)
    assert all(isinstance(n, str) for n in result)
