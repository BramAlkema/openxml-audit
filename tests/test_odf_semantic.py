"""Tests for ODF semantic-core rules and stable IDs."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.odf import OdfValidator, get_odf_semantic_rules


def _semantic_validator() -> OdfValidator:
    return OdfValidator(schema_validation=False, semantic_validation=True)


def test_rule_registry_exposes_stable_unique_ids() -> None:
    rules = get_odf_semantic_rules()
    ids = [rule.id for rule in rules]
    assert ids
    assert len(ids) == len(set(ids))
    assert all(rule.id.startswith("ODFSEM") for rule in rules)


def test_text_style_reference_with_styles_is_valid(
    minimal_odt_text_style_reference_with_styles: Path,
) -> None:
    result = _semantic_validator().validate(minimal_odt_text_style_reference_with_styles)
    assert result.is_valid


def test_presentation_master_page_resolved_is_valid(
    minimal_odp_master_page_resolved: Path,
) -> None:
    result = _semantic_validator().validate(minimal_odp_master_page_resolved)
    assert result.is_valid


def test_manifest_key_part_media_type_mismatch_reports_rule_id(
    odf_manifest_content_bad_media_type: Path,
) -> None:
    result = _semantic_validator().validate(odf_manifest_content_bad_media_type)
    assert not result.is_valid
    assert any(error.id == "ODFSEMMAN001" for error in result.errors)


def test_text_style_reference_without_styles_reports_rule_id(
    odf_text_style_reference_missing_styles: Path,
) -> None:
    result = _semantic_validator().validate(odf_text_style_reference_missing_styles)
    assert not result.is_valid
    assert any(error.id == "ODFSEMTXT001" for error in result.errors)


def test_spreadsheet_duplicate_table_names_report_rule_id(
    odf_spreadsheet_duplicate_table_names: Path,
) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_duplicate_table_names)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSS001" for error in result.errors)


def test_presentation_duplicate_page_names_report_rule_id(
    odf_presentation_duplicate_page_names: Path,
) -> None:
    result = _semantic_validator().validate(odf_presentation_duplicate_page_names)
    assert not result.is_valid
    assert any(error.id == "ODFSEMPRES001" for error in result.errors)


def test_presentation_master_page_missing_styles_reports_rule_id(
    odf_presentation_master_page_missing_styles: Path,
) -> None:
    result = _semantic_validator().validate(odf_presentation_master_page_missing_styles)
    assert not result.is_valid
    assert any(error.id == "ODFSEMREF001" for error in result.errors)


def test_presentation_master_page_unresolved_reports_rule_id(
    odf_presentation_master_page_unresolved: Path,
) -> None:
    result = _semantic_validator().validate(odf_presentation_master_page_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMREF002" for error in result.errors)


def test_content_root_mismatch_uses_stable_rule_id(
    odf_content_root_mismatch: Path,
) -> None:
    result = _semantic_validator().validate(odf_content_root_mismatch)
    assert not result.is_valid
    assert any(error.id == "ODFSEM001" for error in result.errors)


def test_content_body_mismatch_uses_stable_rule_id(
    odf_content_body_mismatch: Path,
) -> None:
    result = _semantic_validator().validate(odf_content_body_mismatch)
    assert not result.is_valid
    assert any(error.id == "ODFSEM004" for error in result.errors)


# ── ODFSEM005: meta.xml must contain office:meta ─────────────────────


def test_meta_missing_office_meta(odf_meta_missing_office_meta: Path) -> None:
    result = _semantic_validator().validate(odf_meta_missing_office_meta)
    assert not result.is_valid
    assert any(error.id == "ODFSEM005" for error in result.errors)


def test_valid_meta_passes(odf_valid_meta: Path) -> None:
    result = _semantic_validator().validate(odf_valid_meta)
    assert result.is_valid


# ── ODFSEM006: settings.xml must contain office:settings ─────────────


def test_settings_missing_office_settings(odf_settings_missing_office_settings: Path) -> None:
    result = _semantic_validator().validate(odf_settings_missing_office_settings)
    assert not result.is_valid
    assert any(error.id == "ODFSEM006" for error in result.errors)


# ── ODFSEMSTYLE001: font-face must have svg:font-family ─────────────


def test_font_face_missing_svg_family(odf_font_face_missing_svg_family: Path) -> None:
    result = _semantic_validator().validate(odf_font_face_missing_svg_family)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSTYLE001" for error in result.errors)


def test_valid_font_face_passes(odf_valid_font_face: Path) -> None:
    result = _semantic_validator().validate(odf_valid_font_face)
    assert result.is_valid


# ── ODFSEMSTYLE002: parent style must resolve ────────────────────────


def test_style_parent_unresolved(odf_style_parent_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_style_parent_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSTYLE002" for error in result.errors)


def test_valid_style_parent_passes(odf_valid_style_parent: Path) -> None:
    result = _semantic_validator().validate(odf_valid_style_parent)
    assert result.is_valid


# ── ODFSEMSTYLE003: data style must resolve ──────────────────────────


def test_style_data_style_unresolved(odf_style_data_style_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_style_data_style_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSTYLE003" for error in result.errors)


# ── ODFSEMSTYLE004: list style must resolve ──────────────────────────


def test_style_list_style_unresolved(odf_style_list_style_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_style_list_style_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSTYLE004" for error in result.errors)


# ── ODFSEMSTYLE005: master page layout must resolve ──────────────────


def test_master_page_layout_unresolved(odf_master_page_layout_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_master_page_layout_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSTYLE005" for error in result.errors)


# ── ODFSEMTXT002: text list style must resolve ───────────────────────


def test_text_list_style_unresolved(odf_text_list_style_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_text_list_style_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMTXT002" for error in result.errors)


# ── ODFSEMTXT003: bookmark ref must resolve ──────────────────────────


def test_text_bookmark_ref_unresolved(odf_text_bookmark_ref_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_text_bookmark_ref_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMTXT003" for error in result.errors)


def test_valid_bookmark_passes(odf_valid_bookmark: Path) -> None:
    result = _semantic_validator().validate(odf_valid_bookmark)
    assert result.is_valid


# ── ODFSEMSS002: named range table must exist ────────────────────────


def test_spreadsheet_named_range_bad_table(odf_spreadsheet_named_range_bad_table: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_named_range_bad_table)
    assert not result.is_valid
    assert any(error.id == "ODFSEMSS002" for error in result.errors)


def test_valid_named_range_passes(odf_valid_named_range: Path) -> None:
    result = _semantic_validator().validate(odf_valid_named_range)
    assert result.is_valid


# ── ODFSEMSS003: row cells must not exceed columns ───────────────────


def test_spreadsheet_column_overflow(odf_spreadsheet_column_overflow: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_column_overflow)
    # This is a WARNING, so the file is still "valid" (no ERROR severity)
    sem_warnings = [e for e in result.errors if e.id == "ODFSEMSS003"]
    assert sem_warnings


# ── ODFSEMPRES002: presentation must have pages ──────────────────────


def test_presentation_no_pages(odf_presentation_no_pages: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_no_pages)
    assert not result.is_valid
    assert any(error.id == "ODFSEMPRES002" for error in result.errors)


# ── ODFSEMPRES003: presentation page layout must resolve ─────────────


def test_presentation_page_layout_unresolved(odf_presentation_page_layout_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_page_layout_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMPRES003" for error in result.errors)


# ── ODFSEMREF004: embedded object href must resolve ──────────────────


def test_embedded_object_unresolved(odf_embedded_object_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_embedded_object_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMREF004" for error in result.errors)


# ── ODFSEMREF005: image href must resolve ────────────────────────────


def test_image_href_unresolved(odf_image_href_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_image_href_unresolved)
    assert not result.is_valid
    assert any(error.id == "ODFSEMREF005" for error in result.errors)


# ── ODFSEMMETA001: document statistics must be valid ─────────────────


def test_meta_bad_statistics(odf_meta_bad_statistics: Path) -> None:
    result = _semantic_validator().validate(odf_meta_bad_statistics)
    assert not result.is_valid
    assert any(error.id == "ODFSEMMETA001" for error in result.errors)
