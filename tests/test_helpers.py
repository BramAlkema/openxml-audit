"""Tests for integration helpers."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from openxml_audit import FileFormat
from openxml_audit.helpers import (
    require_valid_pptx,
    validate_on_save,
    validation_context,
)


class TestValidationContext:
    """Tests for validation_context context manager."""

    def test_basic_usage(self, minimal_pptx: Path) -> None:
        """Test basic context manager usage."""
        with validation_context() as validator:
            result = validator.validate(minimal_pptx)
            assert hasattr(result, "is_valid")

    def test_custom_file_format(self, minimal_pptx: Path) -> None:
        """Test with custom file format."""
        with validation_context(file_format=FileFormat.OFFICE_2007) as validator:
            assert validator.file_format == FileFormat.OFFICE_2007

    def test_custom_max_errors(self, minimal_pptx: Path) -> None:
        """Test with custom max_errors."""
        with validation_context(max_errors=10) as validator:
            assert validator.max_errors == 10

    def test_raise_on_invalid(self, not_a_zip: Path) -> None:
        """Test raise_on_invalid option."""
        with validation_context(raise_on_invalid=True) as validator:
            with pytest.raises(ValueError) as exc_info:
                validator.validate(not_a_zip)

            assert "Invalid PPTX" in str(exc_info.value)

    def test_no_raise_on_invalid(self, not_a_zip: Path) -> None:
        """Test with raise_on_invalid=False (default)."""
        with validation_context(raise_on_invalid=False) as validator:
            # Should not raise, just return result
            result = validator.validate(not_a_zip)
            assert not result.is_valid


class TestValidateOnSaveDecorator:
    """Tests for validate_on_save decorator."""

    def test_decorator_preserves_function(self) -> None:
        """Test decorator preserves function metadata."""

        @validate_on_save()
        def my_function(output_path: str) -> str:
            """My docstring."""
            return "result"

        assert my_function.__name__ == "my_function"
        assert my_function.__doc__ == "My docstring."

    def test_decorator_calls_function(self, tmp_path: Path) -> None:
        """Test decorator still calls wrapped function."""
        called = []

        @validate_on_save()
        def create_something(output_path: str) -> None:
            called.append(output_path)

        output = tmp_path / "output.pptx"
        create_something(str(output))

        assert str(output) in called

    def test_decorator_validates_output(self, tmp_path: Path, minimal_pptx: Path) -> None:
        """Test decorator validates created PPTX."""

        @validate_on_save(raise_on_invalid=True)
        def create_valid_pptx(output_path: str) -> None:
            Path(output_path).write_bytes(minimal_pptx.read_bytes())

        output = tmp_path / "valid_output.pptx"
        # Should not raise for valid PPTX
        create_valid_pptx(str(output))

    def test_decorator_raises_for_invalid(self, tmp_path: Path) -> None:
        """Test decorator raises for invalid output."""

        @validate_on_save(raise_on_invalid=True)
        def create_invalid_pptx(output_path: str) -> None:
            # Create invalid content
            Path(output_path).write_text("not a valid pptx")

        output = tmp_path / "invalid_output.pptx"

        with pytest.raises(ValueError) as exc_info:
            create_invalid_pptx(str(output))

        assert "invalid" in str(exc_info.value).lower()


class TestRequireValidPptxDecorator:
    """Tests for require_valid_pptx decorator."""

    def test_decorator_preserves_function(self) -> None:
        """Test decorator preserves function metadata."""

        @require_valid_pptx()
        def my_function(input_path: str) -> str:
            """Process a PPTX."""
            return "done"

        assert my_function.__name__ == "my_function"
        assert my_function.__doc__ == "Process a PPTX."

    def test_valid_input_passes(self, minimal_pptx: Path) -> None:
        """Test valid input allows function to run."""
        result_holder = []

        @require_valid_pptx()
        def process_pptx(input_path: str) -> str:
            result_holder.append("executed")
            return "processed"

        result = process_pptx(str(minimal_pptx))

        assert result == "processed"
        assert "executed" in result_holder

    def test_invalid_input_raises(self, not_a_zip: Path) -> None:
        """Test invalid input raises before function runs."""
        result_holder = []

        @require_valid_pptx()
        def process_pptx(input_path: str) -> str:
            result_holder.append("executed")
            return "processed"

        with pytest.raises(ValueError) as exc_info:
            process_pptx(str(not_a_zip))

        # Function should not have been called
        assert "executed" not in result_holder
        assert "invalid" in str(exc_info.value).lower()
