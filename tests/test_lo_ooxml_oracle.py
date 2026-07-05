"""Tests for the LibreOffice OOXML roundtrip oracle (Spec 036)."""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.lo_ooxml_repair_oracle import (  # noqa: E402
    _canonical_filter,
    _detect_format,
    _slide_parts_for_location,
    survey_feature_survival,
)
from oracle.odf_window import SofficeNotFoundError, find_soffice  # noqa: E402

from openxml_audit.evidence import EvidenceTier  # noqa: E402
from openxml_audit.pptx.oracle_deck_scaffold import scaffold_root  # noqa: E402
from openxml_audit.reference import build_reference_document  # noqa: E402
from openxml_audit.reference.feature_probes import (  # noqa: E402
    PPTX_FEATURE_SIGNATURES,
    probe_slide,
    probeable_keys,
)
from openxml_audit.reference.pptx_slides import (  # noqa: E402
    build_entrance_fade_slide,
    build_entrance_wipe_slide,
)


def _soffice_available() -> bool:
    try:
        find_soffice()
    except SofficeNotFoundError:
        return False
    return True


class TestFormatRouting:
    def test_detect_format(self):
        assert _detect_format(Path("a.docx")) == "docx"
        assert _detect_format(Path("a.XLSX")) == "xlsx"
        with pytest.raises(ValueError, match="unsupported OOXML extension"):
            _detect_format(Path("a.odt"))

    def test_canonical_filters(self):
        pptx = _canonical_filter("pptx")
        assert pptx("ppt/presentation.xml")
        assert pptx("ppt/slides/slide3.xml")
        assert not pptx("ppt/slides/_rels/slide3.xml.rels")
        assert not pptx("ppt/theme/theme1.xml")
        xlsx = _canonical_filter("xlsx")
        assert xlsx("xl/worksheets/sheet1.xml")
        assert xlsx("xl/sharedStrings.xml")
        assert not xlsx("docProps/core.xml")
        docx = _canonical_filter("docx")
        assert docx("word/document.xml")
        assert not docx("word/theme/theme1.xml")


class TestSlideLocations:
    def test_single_and_span(self):
        assert _slide_parts_for_location("slide 2") == ["ppt/slides/slide2.xml"]
        assert _slide_parts_for_location("slides 7-8") == [
            "ppt/slides/slide7.xml",
            "ppt/slides/slide8.xml",
        ]
        assert _slide_parts_for_location("section 1") == []


class TestFeatureProbes:
    def test_every_registered_pptx_capability_is_probeable(self):
        from openxml_audit.pptx.capabilities import list_capability_findings

        registered = {finding.key for finding in list_capability_findings()}
        assert registered == set(probeable_keys())

    def test_generated_entrance_slides_match_their_own_signatures(self):
        fade = build_entrance_fade_slide()
        wipe = build_entrance_wipe_slide()
        assert probe_slide(fade, "pptx.anim.effect.entr.fade")
        assert not probe_slide(fade, "pptx.anim.effect.entr.wipe")
        assert probe_slide(wipe, "pptx.anim.effect.entr.wipe")
        assert not probe_slide(wipe, "pptx.anim.effect.entr.fade")

    def test_scaffold_probe_slides_match_their_signatures(self):
        root = scaffold_root("timing_oracle")

        def slide(number: int) -> bytes:
            return (root / "ppt" / "slides" / f"slide{number}.xml").read_bytes()

        assert probe_slide(slide(2), "pptx.timing.end-condition.time-offset")
        assert not probe_slide(slide(2), "pptx.timing.end-condition.click")
        assert probe_slide(slide(3), "pptx.timing.end-condition.click")
        assert probe_slide(slide(4), "pptx.timing.repeat-duration")
        assert probe_slide(slide(5), "pptx.timing.restart")
        # The title slide carries no timing at all.
        for key in PPTX_FEATURE_SIGNATURES:
            assert not probe_slide(slide(1), key)

    def test_restart_signature_ignores_tm_root(self):
        # Every deck's tmRoot carries restart="never"; the signature
        # must not count it as feature survival.
        root = scaffold_root("timing_oracle")
        slide2 = (root / "ppt" / "slides" / "slide2.xml").read_bytes()
        assert b'restart="never"' in slide2
        assert not probe_slide(slide2, "pptx.timing.restart")


class TestSurvivalSurvey:
    def test_pristine_reference_deck_reports_full_survival(self, tmp_path):
        result = build_reference_document("pptx", tmp_path, minimum_tier=EvidenceTier.LOADABLE)
        survivals = survey_feature_survival(result.document_path, result.manifest)
        assert len(survivals) == len(result.included_keys)
        for survival in survivals:
            assert survival.slide_parts_present, survival.key
            assert survival.signature_present is True, survival.key

    def test_missing_slide_reported_as_not_present(self, tmp_path):
        result = build_reference_document("pptx", tmp_path, minimum_tier=EvidenceTier.LOADABLE)
        stripped = tmp_path / "stripped.pptx"
        with (
            zipfile.ZipFile(result.document_path) as source,
            zipfile.ZipFile(stripped, "w") as target,
        ):
            for info in source.infolist():
                if info.filename == "ppt/slides/slide2.xml":
                    continue
                target.writestr(info, source.read(info.filename))
        survivals = {s.key: s for s in survey_feature_survival(stripped, result.manifest)}
        fade = survivals["pptx.anim.effect.entr.fade"]
        assert not fade.slide_parts_present
        assert fade.signature_present is None


@pytest.mark.skipif(not _soffice_available(), reason="soffice not installed")
def test_soffice_roundtrip_smoke(tmp_path):
    from oracle.lo_ooxml_repair_oracle import observe

    result = build_reference_document("docx", tmp_path, minimum_tier=EvidenceTier.LOADABLE)
    observation = observe(result.document_path, tmp_path / "work")
    assert observation.outcome in {"preserved", "repaired"}
    assert observation.target_format == "docx"
