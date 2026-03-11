"""Tests for the pytest plugin."""

from __future__ import annotations

from pathlib import Path

import pytest


def test_assert_valid_pptx_passes(assert_valid_pptx, minimal_pptx: Path) -> None:
    """Valid PPTX should not raise."""
    assert_valid_pptx(minimal_pptx)


def test_assert_valid_pptx_fails_on_invalid(assert_valid_pptx, not_a_zip: Path) -> None:
    """Invalid file should fail the test."""
    with pytest.raises(pytest.fail.Exception, match="Validation failed"):
        assert_valid_pptx(not_a_zip)


def test_openxml_validator_fixture(openxml_validator) -> None:
    """Validator fixture should be an OpenXmlValidator instance."""
    from openxml_audit.validator import OpenXmlValidator

    assert isinstance(openxml_validator, OpenXmlValidator)


def test_assert_valid_docx_fixture_exists(assert_valid_docx) -> None:
    """assert_valid_docx fixture should be callable."""
    assert callable(assert_valid_docx)


def test_assert_valid_xlsx_fixture_exists(assert_valid_xlsx) -> None:
    """assert_valid_xlsx fixture should be callable."""
    assert callable(assert_valid_xlsx)
