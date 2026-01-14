"""ODF validator skeleton."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.errors import (
    FileFormat,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
)
from openxml_audit.odf.package import OdfPackage


class OdfValidator:
    """Validate ODF packages using manifest + schema rules."""

    def __init__(self, file_format: FileFormat = FileFormat.ODF_1_3, strict: bool = True):
        self._file_format = file_format
        self._strict = strict

    def validate(self, path: str | Path) -> ValidationResult:
        errors: list[ValidationError] = []

        try:
            with OdfPackage(path) as package:
                errors.extend(package.validate_structure())
        except Exception as exc:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,  # type: ignore[name-defined]
                    description=str(exc),
                )
            )

        return ValidationResult(
            is_valid=not errors,
            errors=errors,
            file_path=str(path),
            file_format=self._file_format,
        )
