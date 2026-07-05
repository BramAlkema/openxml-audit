"""Tests for the canonical reference document pipeline (Spec 034)."""

from __future__ import annotations

import json
import zipfile

import pytest
from lxml import etree

from openxml_audit.evidence import EvidenceTier
from openxml_audit.pptx.oracle_deck_scaffold import materialize_scaffold_package
from openxml_audit.reference import (
    TIER_ORDER,
    build_reference_document,
    collect_ledger,
    qualifies_at,
    tier_rank,
)
from openxml_audit.reference.__main__ import main as reference_main
from openxml_audit.reference.emitters import PPTX_SLIDE_SOURCES, has_emitter
from openxml_audit.validator import OpenXmlValidator

NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


# --- ledger -----------------------------------------------------------------


class TestLedger:
    def test_tier_order_matches_adr001_ladder(self):
        assert TIER_ORDER == (
            EvidenceTier.SCHEMA_VALID,
            EvidenceTier.LOADABLE,
            EvidenceTier.ROUNDTRIP_PRESERVED,
            EvidenceTier.SLIDESHOW_VERIFIED,
            EvidenceTier.UI_AUTHORED,
        )
        assert [tier_rank(tier) for tier in TIER_ORDER] == [0, 1, 2, 3, 4]

    def test_collect_ledger_is_sorted_and_format_tagged(self):
        entries = collect_ledger()
        keys = [(entry.format, entry.finding.key) for entry in entries]
        assert keys == sorted(keys)
        assert all(entry.format in ("docx", "pptx", "xlsx") for entry in entries)

    def test_collect_ledger_rejects_unknown_format(self):
        with pytest.raises(ValueError, match="Unknown reference format"):
            collect_ledger(formats=("pptx", "odt"))

    def test_qualifies_at_uses_highest_registered_rank(self):
        entries = {e.finding.key: e.finding for e in collect_ledger()}
        # Registered LOADABLE qualifies at loadable but not roundtrip.
        timing = entries["pptx.timing.repeat-duration"]
        assert qualifies_at(timing, EvidenceTier.LOADABLE)
        assert not qualifies_at(timing, EvidenceTier.ROUNDTRIP_PRESERVED)
        # SLIDESHOW_VERIFIED ranks above roundtrip-preserved.
        fade = entries["pptx.anim.effect.entr.fade"]
        assert qualifies_at(fade, EvidenceTier.ROUNDTRIP_PRESERVED)
        assert not qualifies_at(fade, EvidenceTier.UI_AUTHORED)


# --- builds -----------------------------------------------------------------


@pytest.fixture(scope="module")
def loadable_build(tmp_path_factory):
    out = tmp_path_factory.mktemp("reference-loadable")
    return {
        fmt: build_reference_document(fmt, out, minimum_tier=EvidenceTier.LOADABLE)
        for fmt in ("docx", "pptx", "xlsx")
    }


class TestPptxReference:
    def test_contains_index_plus_feature_slides(self, loadable_build):
        result = loadable_build["pptx"]
        expected_feature_slides = sum(len(PPTX_SLIDE_SOURCES[key]) for key in result.included_keys)
        with zipfile.ZipFile(result.document_path) as zf:
            slides = sorted(
                name
                for name in zf.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            )
            presentation = etree.fromstring(zf.read("ppt/presentation.xml"))
        assert len(slides) == 1 + expected_feature_slides
        sld_ids = presentation.findall(f"{NS_P}sldIdLst/{NS_P}sldId")
        assert len(sld_ids) == len(slides)

    def test_included_and_excluded_sets(self, loadable_build):
        result = loadable_build["pptx"]
        assert result.included_keys == (
            "pptx.anim.effect.entr.fade",
            "pptx.anim.effect.entr.wipe",
            "pptx.timing.end-condition.click",
            "pptx.timing.end-condition.time-offset",
            "pptx.timing.repeat-duration",
            "pptx.timing.restart",
        )
        assert result.excluded == ()

    def test_generated_entrance_slides_carry_registered_structure(self, loadable_build):
        with zipfile.ZipFile(loadable_build["pptx"].document_path) as zf:
            fade = zf.read("ppt/slides/slide2.xml")
            wipe = zf.read("ppt/slides/slide3.xml")
        assert b'transition="in"' in fade and b'filter="fade"' in fade
        assert b'transition="in"' in wipe and b'filter="wipe(up)"' in wipe

    def test_document_passes_own_validator(self, loadable_build):
        result = OpenXmlValidator().validate(loadable_build["pptx"].document_path)
        assert result.error_count == 0

    def test_default_tier_excludes_loadable_only_findings(self, tmp_path):
        result = build_reference_document("pptx", tmp_path)
        # slideshow-verified ranks above roundtrip-preserved, so the two
        # entrance effects qualify at the default tier.
        assert result.included_keys == (
            "pptx.anim.effect.entr.fade",
            "pptx.anim.effect.entr.wipe",
        )
        reasons = dict(result.excluded)
        assert set(reasons.values()) == {"below-minimum-tier"}
        assert reasons["pptx.timing.restart"] == "below-minimum-tier"
        excluded_manifest = {f["key"]: f["reason"] for f in result.manifest["excluded"]}
        assert excluded_manifest == reasons

    def test_build_is_byte_reproducible(self, loadable_build, tmp_path):
        again = build_reference_document("pptx", tmp_path, minimum_tier=EvidenceTier.LOADABLE)
        assert again.document_path.read_bytes() == loadable_build["pptx"].document_path.read_bytes()
        assert again.manifest_path.read_bytes() == loadable_build["pptx"].manifest_path.read_bytes()


class TestDocxXlsxReferences:
    @pytest.mark.parametrize("fmt", ["docx", "xlsx"])
    def test_builds_validate_with_empty_registries(self, loadable_build, fmt):
        result = loadable_build[fmt]
        assert result.included_keys == ()
        validation = OpenXmlValidator().validate(result.document_path)
        assert validation.error_count == 0

    @pytest.mark.parametrize("fmt", ["docx", "xlsx"])
    def test_manifest_declares_empty_ledger(self, loadable_build, fmt):
        manifest = json.loads(loadable_build[fmt].manifest_path.read_text())
        assert manifest["format"] == fmt
        assert manifest["features"] == []
        assert manifest["excluded"] == []


class TestManifest:
    def test_manifest_reproduces_registered_tiers_verbatim(self, loadable_build):
        manifest = loadable_build["pptx"].manifest
        by_key = {feature["key"]: feature for feature in manifest["features"]}
        findings = {e.finding.key: e.finding for e in collect_ledger()}
        for key, feature in by_key.items():
            assert feature["evidence_tiers"] == [
                tier.value for tier in findings[key].evidence_tiers
            ]
            assert feature["included"] is True
            assert feature["location"].startswith(("slide ", "slides "))
        assert manifest["minimum_tier"] == "loadable"
        assert manifest["spec"] == "034-canonical-reference-documents"

    def test_manifest_has_no_exclusions_at_loadable(self, loadable_build):
        assert loadable_build["pptx"].manifest["excluded"] == []

    def test_manifest_locations_track_multi_slide_spans(self, loadable_build):
        manifest = loadable_build["pptx"].manifest
        locations = {f["key"]: f["location"] for f in manifest["features"]}
        assert locations["pptx.anim.effect.entr.fade"] == "slide 2"
        assert locations["pptx.timing.repeat-duration"] == "slide 6"
        assert locations["pptx.timing.restart"] == "slides 7-8"


# --- emitters ---------------------------------------------------------------


class TestEmitters:
    def test_all_bound_sources_resolve(self):
        from openxml_audit.pptx.oracle_deck_scaffold import scaffold_root
        from openxml_audit.reference.emitters import PptxGeneratedSlide

        for sources in PPTX_SLIDE_SOURCES.values():
            for source in sources:
                if isinstance(source, PptxGeneratedSlide):
                    slide_xml = source.builder()
                    etree.fromstring(slide_xml)
                    continue
                slide = (
                    scaffold_root(source.scaffold)
                    / "ppt"
                    / "slides"
                    / f"slide{source.slide_number}.xml"
                )
                assert slide.is_file(), slide

    def test_has_emitter_matches_registry(self):
        assert has_emitter("pptx", "pptx.timing.restart")
        assert has_emitter("pptx", "pptx.anim.effect.entr.fade")
        assert not has_emitter("docx", "anything")


# --- scaffold conformance (regression for the endCondLst ordering fix) -------


def test_timing_oracle_scaffold_is_schema_valid(tmp_path):
    deck = tmp_path / "timing_oracle.pptx"
    materialize_scaffold_package("timing_oracle", deck)
    result = OpenXmlValidator().validate(deck)
    assert result.error_count == 0, [str(e) for e in result.errors]


# --- CLI --------------------------------------------------------------------


class TestCli:
    def test_status_json(self, capsys):
        assert reference_main(["status", "--json"]) == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload["pptx"]["emitter_gaps"] == []
        assert payload["pptx"]["qualifying_at_loadable"] == 6
        assert payload["docx"]["findings"] == 0
        assert payload["xlsx"]["findings"] == 0

    def test_build_smoke(self, tmp_path, capsys):
        assert (
            reference_main(
                [
                    "build",
                    "--format",
                    "pptx",
                    "--minimum-tier",
                    "loadable",
                    "--out",
                    str(tmp_path),
                ]
            )
            == 0
        )
        out = capsys.readouterr().out
        assert "validated" in out
        assert (tmp_path / "reference.pptx").is_file()
        assert (tmp_path / "reference.pptx.manifest.json").is_file()


def test_reference_documents_have_fixed_zip_timestamps(loadable_build):
    with zipfile.ZipFile(loadable_build["pptx"].document_path) as zf:
        assert {info.date_time for info in zf.infolist()} == {(1980, 1, 1, 0, 0, 0)}
