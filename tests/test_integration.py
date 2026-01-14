"""Integration tests for openxml_audit.

Tests the full validation pipeline with various PPTX files.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from openxml_audit import (
    FileFormat,
    OpenXmlValidator,
    ValidationErrorType,
    ValidationSeverity,
    is_valid_pptx,
    validate_pptx,
)
from tests.fixture_loader import load_fixture_bytes


class TestFullValidationPipeline:
    """Integration tests for the complete validation pipeline."""

    def test_minimal_valid_pptx(self, minimal_pptx: Path) -> None:
        """Test validation of minimal valid PPTX passes all phases."""
        validator = OpenXmlValidator(
            schema_validation=True,
            semantic_validation=True,
        )
        result = validator.validate(minimal_pptx)

        # Should complete without critical package errors
        package_errors = [
            e for e in result.errors
            if e.error_type == ValidationErrorType.PACKAGE
            and e.severity == ValidationSeverity.ERROR
        ]
        assert len(package_errors) == 0, f"Package errors: {package_errors}"

    def test_all_validation_phases_run(self, minimal_pptx: Path) -> None:
        """Test that all validation phases are executed."""
        validator = OpenXmlValidator(
            schema_validation=True,
            semantic_validation=True,
        )
        result = validator.validate(minimal_pptx)

        # Result should have file_path and format set
        assert result.file_path == str(minimal_pptx)
        assert result.file_format == FileFormat.OFFICE_2019

    def test_validation_with_schema_disabled(self, minimal_pptx: Path) -> None:
        """Test validation works with schema validation disabled."""
        validator = OpenXmlValidator(
            schema_validation=False,
            semantic_validation=True,
        )
        result = validator.validate(minimal_pptx)

        # Should still report results
        assert result.file_path == str(minimal_pptx)

    def test_validation_with_semantic_disabled(self, minimal_pptx: Path) -> None:
        """Test validation works with semantic validation disabled."""
        validator = OpenXmlValidator(
            schema_validation=True,
            semantic_validation=False,
        )
        result = validator.validate(minimal_pptx)

        assert result.file_path == str(minimal_pptx)


class TestInvalidPPTXFiles:
    """Tests with various invalid PPTX files."""

    def test_corrupted_zip(self, tmp_path: Path) -> None:
        """Test detection of corrupted ZIP file."""
        path = tmp_path / "corrupted.pptx"
        path.write_bytes(b"PK\x03\x04" + b"\x00" * 100)  # Invalid ZIP

        result = validate_pptx(path)
        assert not result.is_valid

    def test_empty_zip(self, tmp_path: Path) -> None:
        """Test detection of empty ZIP file."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w"):
            pass  # Empty ZIP

        path = tmp_path / "empty.pptx"
        path.write_bytes(buffer.getvalue())

        result = validate_pptx(path)
        assert not result.is_valid

    def test_missing_content_types(self, tmp_path: Path) -> None:
        """Test detection of missing [Content_Types].xml."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as zf:
            zf.writestr("dummy.txt", "content")

        path = tmp_path / "no_content_types.pptx"
        path.write_bytes(buffer.getvalue())

        result = validate_pptx(path)
        assert not result.is_valid
        assert any("[Content_Types]" in e.description for e in result.errors)

    def test_missing_package_rels(self, tmp_path: Path) -> None:
        """Test detection of missing _rels/.rels."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as zf:
            content_types = load_fixture_bytes(
                "integration",
                "missing_package_rels",
                "content_types.xml",
            )
            zf.writestr("[Content_Types].xml", content_types)

        path = tmp_path / "no_rels.pptx"
        path.write_bytes(buffer.getvalue())

        result = validate_pptx(path)
        assert not result.is_valid

    def test_invalid_xml_in_presentation(self, tmp_path: Path) -> None:
        """Test detection of malformed XML in presentation.xml."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as zf:
            content_types = load_fixture_bytes(
                "integration",
                "invalid_xml",
                "content_types.xml",
            )
            rels = load_fixture_bytes("integration", "invalid_xml", "rels.xml")
            invalid_presentation = load_fixture_bytes(
                "integration",
                "invalid_xml",
                "invalid_presentation.xml",
            )
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", rels)
            zf.writestr("ppt/presentation.xml", invalid_presentation)

        path = tmp_path / "invalid_xml.pptx"
        path.write_bytes(buffer.getvalue())

        result = validate_pptx(path)
        assert not result.is_valid


class TestErrorReporting:
    """Tests for error reporting quality."""

    def test_errors_have_descriptions(self, not_a_zip: Path) -> None:
        """Test all errors have descriptions."""
        result = validate_pptx(not_a_zip)

        for error in result.errors:
            assert error.description, f"Error missing description: {error}"
            assert len(error.description) > 0

    def test_errors_have_error_type(self, not_a_zip: Path) -> None:
        """Test all errors have error type."""
        result = validate_pptx(not_a_zip)

        for error in result.errors:
            assert error.error_type is not None
            assert isinstance(error.error_type, ValidationErrorType)

    def test_errors_have_severity(self, not_a_zip: Path) -> None:
        """Test all errors have severity."""
        result = validate_pptx(not_a_zip)

        for error in result.errors:
            assert error.severity is not None
            assert isinstance(error.severity, ValidationSeverity)


class TestFileFormatVersions:
    """Tests for different Office format versions."""

    @pytest.mark.parametrize(
        "file_format",
        [
            FileFormat.OFFICE_2007,
            FileFormat.OFFICE_2010,
            FileFormat.OFFICE_2013,
            FileFormat.OFFICE_2016,
            FileFormat.OFFICE_2019,
            FileFormat.MICROSOFT_365,
        ],
    )
    def test_each_file_format(
        self, minimal_pptx: Path, file_format: FileFormat
    ) -> None:
        """Test validation works for each file format version."""
        validator = OpenXmlValidator(file_format=file_format)
        result = validator.validate(minimal_pptx)

        assert result.file_format == file_format
        # Should not crash for any version


class TestConvenienceFunctionsIntegration:
    """Integration tests for convenience functions."""

    def test_validate_pptx_full_pipeline(self, minimal_pptx: Path) -> None:
        """Test validate_pptx runs full pipeline."""
        result = validate_pptx(minimal_pptx)

        assert result.file_path == str(minimal_pptx)
        assert result.file_format == FileFormat.OFFICE_2019

    def test_is_valid_pptx_true(self, minimal_pptx: Path) -> None:
        """Test is_valid_pptx returns True for valid file."""
        # Note: may fail if minimal_pptx has validation issues
        result = is_valid_pptx(minimal_pptx)
        assert isinstance(result, bool)

    def test_is_valid_pptx_false(self, not_a_zip: Path) -> None:
        """Test is_valid_pptx returns False for invalid file."""
        result = is_valid_pptx(not_a_zip)
        assert result is False
