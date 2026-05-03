"""Regression tests for known GSuite import failure modes.

Each test pairs a known-bad fixture (triggers the failure) with the
known-good counterpart (failure mode fixed). The bad-side test
asserts the oracle correctly captures the failure as
`convert_failed`; the good-side asserts a clean `lossy_conversion`.

All tests skip cleanly when GSuite oracle credentials aren't
configured — the fixtures alone document the failure mode in the
corpus, the live tests verify our oracle observes it correctly.

Fixtures live in `data/corpus/gsuite_known_failures/<format>/`. See
that directory's README.md for the bisection evidence behind each
case.
"""

from __future__ import annotations

import os
import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.gsuite_roundtrip import observe  # noqa: E402

FIXTURES = REPO_ROOT / "data" / "corpus" / "gsuite_known_failures" / "pptx"


def _gsuite_configured() -> bool:
    creds_default = (
        Path.home() / ".config" / "openxml-audit" / "google_service_account.json"
    )
    has_creds = bool(os.environ.get("GSUITE_ORACLE_CREDS")) or creds_default.exists()
    return (
        has_creds
        and bool(os.environ.get("GSUITE_ORACLE_SUBJECT"))
        and bool(os.environ.get("GSUITE_ORACLE_FOLDER_ID"))
    )


REQUIRES_GSUITE = pytest.mark.skipif(
    not _gsuite_configured(),
    reason="GSuite oracle requires creds + GSUITE_ORACLE_SUBJECT + "
           "GSUITE_ORACLE_FOLDER_ID",
)


# --- Fixture-shape sanity checks (always-on) -------------------------------


def test_bad_fixture_has_stroke_zero_attribute():
    """The known-bad fixture must contain the failure-triggering construct,
    or the regression test means nothing."""
    bad = FIXTURES / "bee-with-stroke-zero.pptx"
    assert bad.exists(), f"missing fixture: {bad}"
    with zipfile.ZipFile(bad) as z:
        slide = z.read("ppt/slides/slide1.xml").decode("utf-8")
    actual = slide.count('stroke="0"')
    assert actual == 33, f"expected 33 stroke=\"0\" attrs; got {actual}"


def test_good_fixture_does_not_have_stroke_zero_attribute():
    """The known-good fixture must have the construct removed."""
    good = FIXTURES / "bee-without-stroke-zero.pptx"
    assert good.exists(), f"missing fixture: {good}"
    with zipfile.ZipFile(good) as z:
        slide = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert 'stroke="0"' not in slide, "good fixture should not contain stroke=\"0\""


def test_fixtures_have_same_shape_count():
    """Sanity: both fixtures should have the same 33 shapes — the only
    intentional difference is the stroke attribute."""
    counts = {}
    for name in ("bee-with-stroke-zero.pptx", "bee-without-stroke-zero.pptx"):
        with zipfile.ZipFile(FIXTURES / name) as z:
            slide = z.read("ppt/slides/slide1.xml").decode("utf-8")
        counts[name] = slide.count("<p:sp>")
    assert counts["bee-with-stroke-zero.pptx"] == 33
    assert counts["bee-without-stroke-zero.pptx"] == 33


# --- Live GSuite regression tests (gated) ----------------------------------


@REQUIRES_GSUITE
def test_stroke_zero_breaks_gsuite_import():
    """svg2ooxml's old `<a:path stroke="0">` emission causes Google
    Slides' `files.copy` to return HTTP 500.

    If this test ever flips to `lossy_conversion` it means Google
    fixed the import bug server-side — interesting signal worth
    investigating, but not a regression in our code. Update the
    fixture's README.md and either remove this test or mark it
    xfail with a date stamp."""
    bad = FIXTURES / "bee-with-stroke-zero.pptx"
    obs = observe(bad)
    assert obs.outcome == "convert_failed", (
        f"expected convert_failed (HTTP 500 from Google's importer); "
        f"got {obs.outcome}. Notes: {obs.notes}"
    )
    assert any("500" in note or "Internal Error" in note for note in obs.notes), (
        f"expected HTTP 500 in notes; got {obs.notes}"
    )


@REQUIRES_GSUITE
def test_stroke_zero_removed_imports_cleanly():
    """The fixed bee (svg2ooxml patched to omit `stroke="0"`) imports
    cleanly through GSuite — same 33 shapes, just the bad attribute
    removed."""
    good = FIXTURES / "bee-without-stroke-zero.pptx"
    obs = observe(good)
    assert obs.outcome == "lossy_conversion", (
        f"expected lossy_conversion; got {obs.outcome}. Notes: {obs.notes}"
    )
    # Empirically these two classes always fire on a GSuite roundtrip.
    assert "metadata_churn" in obs.loss
    assert "structural_normalization" in obs.loss
