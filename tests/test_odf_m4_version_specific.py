"""Tests for M4 version-specific ODF semantic rules."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.odf import OdfValidator
from openxml_audit.odf.constraints.base import parse_version_tuple


def _semantic_validator() -> OdfValidator:
    return OdfValidator(schema_validation=False, semantic_validation=True)


# ── Infrastructure tests ─────────────────────────────────────────────────


def test_parse_version_tuple_basic() -> None:
    assert parse_version_tuple("1.3") == (1, 3)
    assert parse_version_tuple("1.2") == (1, 2)
    assert parse_version_tuple("1.4") == (1, 4)


def test_parse_version_tuple_none() -> None:
    assert parse_version_tuple(None) is None
    assert parse_version_tuple("") is None
    assert parse_version_tuple("  ") is None


def test_parse_version_tuple_with_extras() -> None:
    assert parse_version_tuple("1.3+csd01") == (1, 3)


def test_version_gating_skips_constraint() -> None:
    """Constraints with min_version should be skipped for older docs."""
    from openxml_audit.errors import ValidationError, ValidationErrorType
    from openxml_audit.odf.constraints.base import (
        EvaluationContext,
        OdfConstraint,
        OdfSemanticRule,
    )

    class TestConstraint(OdfConstraint):
        @property
        def rule(self) -> OdfSemanticRule:
            return OdfSemanticRule(
                id="TEST001",
                family="test",
                description="Test",
                min_version="1.3",
            )

        def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
            return [
                self._error(
                    rule_id="TEST001",
                    error_type=ValidationErrorType.SEMANTIC,
                    description="should not fire",
                    part_uri="/content.xml",
                )
            ]

    c = TestConstraint()
    assert c.applies_to_version("1.3") is True
    assert c.applies_to_version("1.4") is True
    assert c.applies_to_version("1.2") is False
    assert c.applies_to_version("1.1") is False
    assert c.applies_to_version(None) is True  # unknown version → apply


# ── Version constraint tests ─────────────────────────────────────────────


def test_version_missing_attr(odf_version_missing_attr: Path) -> None:
    result = _semantic_validator().validate(odf_version_missing_attr)
    assert any(e.id == "ODFSEMVER001" for e in result.errors)


def test_version_inconsistent(odf_version_inconsistent: Path) -> None:
    result = _semantic_validator().validate(odf_version_inconsistent)
    assert any(e.id == "ODFSEMVER002" for e in result.errors)


def test_rdf_pre12(odf_version_rdf_pre12: Path) -> None:
    result = _semantic_validator().validate(odf_version_rdf_pre12)
    assert not result.is_valid
    assert any(e.id == "ODFSEMVER003" for e in result.errors)


def test_named_expr_pre12(odf_version_named_expr_pre12: Path) -> None:
    result = _semantic_validator().validate(odf_version_named_expr_pre12)
    assert not result.is_valid
    assert any(e.id == "ODFSEMVER004" for e in result.errors)


def test_signature_pre12(odf_version_signature_pre12: Path) -> None:
    result = _semantic_validator().validate(odf_version_signature_pre12)
    assert not result.is_valid
    assert any(e.id == "ODFSEMVER005" for e in result.errors)


def test_change_tracking_pre12(odf_version_change_tracking_pre12: Path) -> None:
    result = _semantic_validator().validate(odf_version_change_tracking_pre12)
    assert not result.is_valid
    assert any(e.id == "ODFSEMVER006" for e in result.errors)


def test_draw_enhanced_pre13(odf_version_draw_enhanced_pre13: Path) -> None:
    result = _semantic_validator().validate(odf_version_draw_enhanced_pre13)
    assert not result.is_valid
    assert any(e.id == "ODFSEMVER007" for e in result.errors)


def test_anim_iterate_pre13(odf_version_anim_iterate_pre13: Path) -> None:
    result = _semantic_validator().validate(odf_version_anim_iterate_pre13)
    assert not result.is_valid
    assert any(e.id == "ODFSEMVER008" for e in result.errors)


# ── Valid fixtures should still pass ─────────────────────────────────────


def test_minimal_odt_still_valid(minimal_odt: Path) -> None:
    result = _semantic_validator().validate(minimal_odt)
    assert result.is_valid


def test_minimal_odt_v12_still_valid(minimal_odt_v12: Path) -> None:
    result = _semantic_validator().validate(minimal_odt_v12)
    assert result.is_valid
