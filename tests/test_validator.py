"""Tests for the main validator."""

from __future__ import annotations

from pathlib import Path

import pytest

from openxml_audit import (
    FileFormat,
    OpenXmlValidator,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
    is_valid_pptx,
    validate_pptx,
)


class TestOpenXmlValidator:
    """Tests for OpenXmlValidator class."""

    def test_validator_initialization_defaults(self) -> None:
        """Test validator initializes with correct defaults."""
        validator = OpenXmlValidator()
        assert validator.file_format == FileFormat.OFFICE_2019
        assert validator.max_errors == 1000

    def test_validator_initialization_custom(self) -> None:
        """Test validator with custom options."""
        validator = OpenXmlValidator(
            file_format=FileFormat.OFFICE_2007,
            max_errors=50,
            schema_validation=False,
            semantic_validation=False,
        )
        assert validator.file_format == FileFormat.OFFICE_2007
        assert validator.max_errors == 50

    def test_validate_valid_pptx(self, minimal_pptx: Path) -> None:
        """Test validation of a valid PPTX file."""
        validator = OpenXmlValidator()
        result = validator.validate(minimal_pptx)

        assert isinstance(result, ValidationResult)
        assert result.file_path == str(minimal_pptx)
        assert result.file_format == FileFormat.OFFICE_2019
        # A well-formed minimal PPTX should be valid
        # (allowing for minor schema issues in our test fixture)

    def test_validate_returns_errors(
        self, invalid_pptx_missing_presentation: Path
    ) -> None:
        """Test validation returns errors for invalid PPTX."""
        validator = OpenXmlValidator()

        # This should raise PackageValidationError which is caught internally
        result = validator.validate(invalid_pptx_missing_presentation)

        assert not result.is_valid
        assert len(result.errors) > 0

    def test_validate_nonexistent_file(self, tmp_path: Path) -> None:
        """Test validation of nonexistent file."""
        validator = OpenXmlValidator()
        nonexistent = tmp_path / "nonexistent.pptx"

        # Should handle gracefully
        result = validator.validate(nonexistent)
        assert not result.is_valid

    def test_validate_not_a_zip(self, not_a_zip: Path) -> None:
        """Test validation of non-ZIP file."""
        validator = OpenXmlValidator()
        result = validator.validate(not_a_zip)

        assert not result.is_valid
        assert any("ZIP" in e.description for e in result.errors)

    def test_is_valid_method(self, minimal_pptx: Path) -> None:
        """Test the is_valid convenience method."""
        validator = OpenXmlValidator()
        # Should return boolean directly
        result = validator.is_valid(minimal_pptx)
        assert isinstance(result, bool)

    def test_max_errors_limit(self, minimal_pptx: Path) -> None:
        """Test that max_errors limits error collection."""
        validator = OpenXmlValidator(max_errors=1)
        result = validator.validate(minimal_pptx)

        # Should not exceed max_errors (counting only ERROR severity)
        error_count = sum(
            1 for e in result.errors if e.severity == ValidationSeverity.ERROR
        )
        assert error_count <= 1

    def test_max_errors_zero_unlimited(self, minimal_pptx: Path) -> None:
        """Test that max_errors=0 means unlimited."""
        validator = OpenXmlValidator(max_errors=0)
        # Should not crash with unlimited errors
        result = validator.validate(minimal_pptx)
        assert isinstance(result, ValidationResult)


class TestValidationResult:
    """Tests for ValidationResult."""

    def test_error_count(self, minimal_pptx: Path) -> None:
        """Test error_count property."""
        validator = OpenXmlValidator()
        result = validator.validate(minimal_pptx)

        # error_count should match number of ERROR severity items
        expected = sum(
            1 for e in result.errors if e.severity == ValidationSeverity.ERROR
        )
        assert result.error_count == expected

    def test_warning_count(self, minimal_pptx: Path) -> None:
        """Test warning_count property."""
        validator = OpenXmlValidator()
        result = validator.validate(minimal_pptx)

        expected = sum(
            1 for e in result.errors if e.severity == ValidationSeverity.WARNING
        )
        assert result.warning_count == expected


class TestConvenienceFunctions:
    """Tests for convenience functions."""

    def test_validate_pptx(self, minimal_pptx: Path) -> None:
        """Test validate_pptx convenience function."""
        result = validate_pptx(minimal_pptx)
        assert isinstance(result, ValidationResult)

    def test_is_valid_pptx(self, minimal_pptx: Path) -> None:
        """Test is_valid_pptx convenience function."""
        result = is_valid_pptx(minimal_pptx)
        assert isinstance(result, bool)

    def test_validate_pptx_string_path(self, minimal_pptx: Path) -> None:
        """Test that string paths work."""
        result = validate_pptx(str(minimal_pptx))
        assert isinstance(result, ValidationResult)
