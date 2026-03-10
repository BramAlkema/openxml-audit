"""Tests for M3 style inheritance chain ODF semantic rules."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.odf import OdfValidator


def _semantic_validator() -> OdfValidator:
    return OdfValidator(schema_validation=False, semantic_validation=True)


# ── Style chain constraints ──────────────────────────────────────────────


def test_style_cycle(odf_style_cycle: Path) -> None:
    result = _semantic_validator().validate(odf_style_cycle)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN001" for e in result.errors)


def test_orphaned_auto_style(odf_style_orphaned_auto: Path) -> None:
    result = _semantic_validator().validate(odf_style_orphaned_auto)
    assert any(e.id == "ODFSEMCHAIN002" for e in result.errors)


def test_no_default_style(odf_style_no_default: Path) -> None:
    result = _semantic_validator().validate(odf_style_no_default)
    assert any(e.id == "ODFSEMCHAIN003" for e in result.errors)


def test_style_map_target_unresolved(odf_style_map_target_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_style_map_target_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN004" for e in result.errors)


def test_master_page_no_layout(odf_style_master_page_no_layout: Path) -> None:
    result = _semantic_validator().validate(odf_style_master_page_no_layout)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN005" for e in result.errors)


def test_style_family_mismatch(odf_style_family_mismatch: Path) -> None:
    result = _semantic_validator().validate(odf_style_family_mismatch)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN006" for e in result.errors)


def test_style_name_empty(odf_style_name_empty: Path) -> None:
    result = _semantic_validator().validate(odf_style_name_empty)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN007" for e in result.errors)


def test_style_duplicate_name(odf_style_duplicate_name: Path) -> None:
    result = _semantic_validator().validate(odf_style_duplicate_name)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN008" for e in result.errors)


def test_deep_inheritance(odf_style_deep_inheritance: Path) -> None:
    result = _semantic_validator().validate(odf_style_deep_inheritance)
    assert any(e.id == "ODFSEMCHAIN009" for e in result.errors)


def test_next_style_unresolved(odf_style_next_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_style_next_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN010" for e in result.errors)


def test_master_page_next_unresolved(odf_style_master_next_unresolved: Path) -> None:
    result = _semantic_validator().validate(odf_style_master_next_unresolved)
    assert not result.is_valid
    assert any(e.id == "ODFSEMCHAIN011" for e in result.errors)


def test_font_name_undeclared(odf_style_font_name_undeclared: Path) -> None:
    result = _semantic_validator().validate(odf_style_font_name_undeclared)
    assert any(e.id == "ODFSEMCHAIN012" for e in result.errors)


# ── Existing valid fixtures should still pass ────────────────────────────


def test_minimal_odt_still_valid(minimal_odt: Path) -> None:
    result = _semantic_validator().validate(minimal_odt)
    assert result.is_valid


def test_minimal_odt_with_styles_still_valid(minimal_odt_with_styles: Path) -> None:
    result = _semantic_validator().validate(minimal_odt_with_styles)
    assert result.is_valid
