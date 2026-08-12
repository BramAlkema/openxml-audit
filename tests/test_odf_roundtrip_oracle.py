"""Tests for the ODF roundtrip oracle (Spec 019, 0.6.3).

Tests are split into:

  - **always-on**: pure-Python harness logic that doesn't require soffice
    (argument parsing, fingerprint computation, JSON shaping, supported
    formats lookup).
  - **soffice-required**: real-file roundtrip smoke. Skipped when no
    soffice binary is on disk so CI without LibreOffice still passes.
"""

from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.odf_repair_oracle import (  # noqa: E402
    PartFingerprint,
    RoundtripObservation,
    _detect_format,
    _fingerprint_parts,
    _to_jsonable,
)
from oracle.odf_window import (  # noqa: E402
    SofficeNotFoundError,
    SofficeRunResult,
    find_soffice,
    roundtrip,
)


def _has_soffice() -> bool:
    try:
        find_soffice()
    except SofficeNotFoundError:
        return False
    return True


REQUIRES_SOFFICE = pytest.mark.skipif(
    not _has_soffice(), reason="soffice not installed; ODF oracle tests skipped"
)


def _build_minimal_odt(target: Path) -> Path:
    """Write a tiny but valid .odt at `target` for tests that don't
    require soffice. The file is hand-built so it doesn't depend on
    TokenMoulds or any external generator."""
    target.parent.mkdir(parents=True, exist_ok=True)
    mimetype = b"application/vnd.oasis.opendocument.text"
    content_xml = (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b"<office:document-content "
        b'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        b'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        b'office:version="1.3">'
        b"<office:body><office:text>"
        b"<text:p>hello</text:p>"
        b"</office:text></office:body>"
        b"</office:document-content>"
    )
    manifest_xml = (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b"<manifest:manifest "
        b'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" '
        b'manifest:version="1.3">'
        b'<manifest:file-entry manifest:full-path="/" '
        b'manifest:media-type="application/vnd.oasis.opendocument.text"/>'
        b'<manifest:file-entry manifest:full-path="content.xml" '
        b'manifest:media-type="text/xml"/>'
        b"</manifest:manifest>"
    )
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(zipfile.ZipInfo("mimetype"), mimetype, compress_type=zipfile.ZIP_STORED)
        zf.writestr("content.xml", content_xml)
        zf.writestr("META-INF/manifest.xml", manifest_xml)
    return target


# -- always-on tests ---------------------------------------------------------


def test_detect_format_known_extensions() -> None:
    assert _detect_format(Path("foo.odt")) == "odt"
    assert _detect_format(Path("foo.ODT")) == "odt"
    assert _detect_format(Path("foo.ods")) == "ods"
    assert _detect_format(Path("foo.odp")) == "odp"
    assert _detect_format(Path("foo.odg")) == "odg"


def test_detect_format_rejects_unsupported() -> None:
    with pytest.raises(ValueError):
        _detect_format(Path("foo.docx"))


def test_fingerprint_parts_skips_missing(tmp_path: Path) -> None:
    odt = _build_minimal_odt(tmp_path / "minimal.odt")
    fps = _fingerprint_parts(odt)
    names = [fp.name for fp in fps]
    # Our minimal builder includes content.xml only (not styles/meta/settings).
    assert names == ["content.xml"]
    assert all(isinstance(fp, PartFingerprint) for fp in fps)
    assert fps[0].size_bytes > 0
    assert len(fps[0].sha256) == 64  # sha256 hex


def test_fingerprint_parts_returns_empty_for_missing_file(tmp_path: Path) -> None:
    fps = _fingerprint_parts(tmp_path / "nope.odt")
    assert fps == []


def test_to_jsonable_summary_counts_outcomes() -> None:
    obs = [
        RoundtripObservation(
            source_relpath="a.odt", target_format="odt", outcome="preserved", duration_seconds=1.0
        ),
        RoundtripObservation(
            source_relpath="b.odt", target_format="odt", outcome="repaired", duration_seconds=2.0
        ),
        RoundtripObservation(
            source_relpath="c.odt", target_format="odt", outcome="crash", duration_seconds=0.5
        ),
    ]
    payload = _to_jsonable(obs)
    assert payload["schema_version"] == 1
    assert payload["summary"] == {
        "total": 3,
        "preserved": 1,
        "repaired": 1,
        "crash": 1,
        "timeout": 0,
        "open_failed": 0,
        "missing_output": 0,
    }
    # Round-trip through json to confirm the shape is JSON-serializable.
    assert json.loads(json.dumps(payload))["summary"]["total"] == 3


def test_soffice_run_result_default_fields() -> None:
    result = SofficeRunResult(
        outcome="ok",
        output_path=Path("/tmp/x.odt"),
        duration_seconds=1.5,
    )
    assert result.return_code is None
    assert result.stderr == ""
    assert result.notes == []


# -- soffice-required tests --------------------------------------------------


@REQUIRES_SOFFICE
def test_roundtrip_minimal_odt_succeeds(tmp_path: Path) -> None:
    """Real soffice call: the minimal hand-built ODT roundtrips clean."""
    in_odt = _build_minimal_odt(tmp_path / "in" / "minimal.odt")
    result = roundtrip(in_odt, tmp_path / "out", target_format="odt", timeout_seconds=120)
    assert result.outcome == "ok", (
        f"expected ok, got {result.outcome}; stderr={result.stderr[:500]}"
    )
    assert result.output_path is not None
    assert result.output_path.exists()
    assert result.duration_seconds > 0


@REQUIRES_SOFFICE
def test_roundtrip_missing_input_returns_structured_error(tmp_path: Path) -> None:
    """Harness returns a `missing_output` result rather than raising on
    missing input."""
    result = roundtrip(
        tmp_path / "nope.odt",
        tmp_path / "out",
        target_format="odt",
        timeout_seconds=10,
    )
    assert result.outcome == "missing_output"
    assert result.output_path is None
    assert any("does not exist" in note for note in result.notes)
