"""Tests for M2 format-specific ODF semantic rules."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.odf import OdfValidator


def _semantic_validator() -> OdfValidator:
    return OdfValidator(schema_validation=False, semantic_validation=True)


# ── ODT rules ────────────────────────────────────────────────────────────


def test_heading_level_skip(odf_text_heading_level_skip: Path) -> None:
    result = _semantic_validator().validate(odf_text_heading_level_skip)
    sem = [e for e in result.errors if e.id == "ODFSEMTXT004"]
    assert sem


def test_note_ref_unresolved(odf_text_note_ref_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_text_note_ref_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT005" for e in result.errors)


def test_section_duplicate_name(odf_text_section_duplicate_name: Path) -> None:
    result = _semantic_validator().validate(odf_text_section_duplicate_name)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT006" for e in result.errors)


def test_tracked_change_dup_id(odf_text_tracked_change_dup_id: Path) -> None:
    result = _semantic_validator().validate(odf_text_tracked_change_dup_id)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT007" for e in result.errors)


def test_sequence_decl_duplicate(odf_text_sequence_decl_duplicate: Path) -> None:
    result = _semantic_validator().validate(odf_text_sequence_decl_duplicate)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT008" for e in result.errors)


def test_variable_decl_duplicate(odf_text_variable_decl_duplicate: Path) -> None:
    result = _semantic_validator().validate(odf_text_variable_decl_duplicate)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT009" for e in result.errors)


def test_variable_get_unresolved(odf_text_variable_get_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_text_variable_get_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT010" for e in result.errors)


def test_user_field_decl_duplicate(odf_text_user_field_decl_duplicate: Path) -> None:
    result = _semantic_validator().validate(odf_text_user_field_decl_duplicate)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT011" for e in result.errors)


def test_user_field_get_unresolved(odf_text_user_field_get_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_text_user_field_get_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT012" for e in result.errors)


def test_text_table_empty(odf_text_table_empty: Path) -> None:
    result = _semantic_validator().validate(odf_text_table_empty)
    assert not result.is_valid
    assert any(e.id == "ODFSEMTXT013" for e in result.errors)


# ── ODS rules ────────────────────────────────────────────────────────────


def test_spreadsheet_no_tables(odf_spreadsheet_no_tables: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_no_tables)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS004" for e in result.errors)


def test_database_range_bad_table(odf_spreadsheet_database_range_bad_table: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_database_range_bad_table)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS005" for e in result.errors)


def test_data_pilot_bad_source(odf_spreadsheet_data_pilot_bad_source: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_data_pilot_bad_source)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS006" for e in result.errors)


def test_validation_duplicate(odf_spreadsheet_validation_duplicate: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_validation_duplicate)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS007" for e in result.errors)


def test_validation_ref_unresolved(odf_spreadsheet_validation_ref_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_validation_ref_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS008" for e in result.errors)


def test_bad_repeat_count(odf_spreadsheet_bad_repeat_count: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_bad_repeat_count)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS009" for e in result.errors)


def test_column_style_unresolved(odf_spreadsheet_column_style_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_column_style_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS010" for e in result.errors)


def test_cell_style_unresolved(odf_spreadsheet_cell_style_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_cell_style_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS011" for e in result.errors)


def test_conditional_style_unresolved(
    odf_spreadsheet_conditional_style_unresolved: Path,
) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_conditional_style_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS012" for e in result.errors)


def test_filter_bad_field(odf_spreadsheet_filter_bad_field: Path) -> None:
    result = _semantic_validator().validate(odf_spreadsheet_filter_bad_field)
    assert not result.is_valid
    assert any(e.id == "ODFSEMSS013" for e in result.errors)


# ── ODP rules ────────────────────────────────────────────────────────────


def test_custom_show_bad_ref(odf_presentation_custom_show_bad_ref: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_custom_show_bad_ref)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES004" for e in result.errors)


def test_duplicate_layer(odf_presentation_duplicate_layer: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_duplicate_layer)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES005" for e in result.errors)


def test_sound_unresolved(odf_presentation_sound_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_sound_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES006" for e in result.errors)


def test_header_decl_duplicate(odf_presentation_header_decl_duplicate: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_header_decl_duplicate)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES007" for e in result.errors)


def test_header_ref_unresolved(odf_presentation_header_ref_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_header_ref_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES008" for e in result.errors)


def test_settings_bad_start(odf_presentation_settings_bad_start: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_settings_bad_start)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES009" for e in result.errors)


def test_animation_bad_target(odf_presentation_animation_bad_target: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_animation_bad_target)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES010" for e in result.errors)


def test_bad_transition_type(odf_presentation_bad_transition_type: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_bad_transition_type)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES011" for e in result.errors)


def test_notes_bad_ref(odf_presentation_notes_bad_ref: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_notes_bad_ref)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES012" for e in result.errors)


def test_bad_class(odf_presentation_bad_class: Path) -> None:
    result = _semantic_validator().validate(odf_presentation_bad_class)
    assert not result.is_valid
    assert any(e.id == "ODFSEMPRES013" for e in result.errors)


# ── Existing valid fixtures should still pass ────────────────────────────


def test_minimal_odt_still_valid(minimal_odt: Path) -> None:
    result = _semantic_validator().validate(minimal_odt)
    assert result.is_valid


def test_minimal_ods_still_valid(minimal_ods: Path) -> None:
    result = _semantic_validator().validate(minimal_ods)
    assert result.is_valid


def test_minimal_odp_still_valid(minimal_odp: Path) -> None:
    result = _semantic_validator().validate(minimal_odp)
    assert result.is_valid
