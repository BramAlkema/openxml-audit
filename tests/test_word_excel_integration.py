"""Integration tests for Word and Excel validators."""

from __future__ import annotations

from pathlib import Path

from openxml_audit import OpenXmlValidator, ValidationErrorType


def test_minimal_docx_valid(minimal_docx: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(minimal_docx)

    assert result.is_valid


def test_docx_missing_body_reports_error(docx_missing_body: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(docx_missing_body)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SCHEMA and "body" in e.description
        for e in result.errors
    )


def test_docx_missing_styles_with_effects_reports_error(
    docx_missing_styles_with_effects: Path,
) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=True)
    result = validator.validate(docx_missing_styles_with_effects)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SEMANTIC
        and "stylesWithEffects" in e.description
        for e in result.errors
    )


def test_minimal_xlsx_valid(minimal_xlsx: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(minimal_xlsx)

    assert result.is_valid


def test_xlsx_missing_sheet_relationship_reports_error(xlsx_missing_sheet_rel: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(xlsx_missing_sheet_rel)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SEMANTIC and "relationship" in e.description
        for e in result.errors
    )


def test_xlsx_missing_shared_strings_reports_error(
    xlsx_missing_shared_strings: Path,
) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=True)
    result = validator.validate(xlsx_missing_shared_strings)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SEMANTIC
        and "sharedStrings" in e.description
        for e in result.errors
    )
