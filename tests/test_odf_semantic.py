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
