"""Shared core utilities for document validators."""

from openxml_audit.core.context import ElementContext, ValidationContext, ValidationStack
from openxml_audit.core.errors import (
    FileFormat,
    PackageValidationError,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.core.package import ZipPackage

__all__ = [
    "ElementContext",
    "ValidationContext",
    "ValidationStack",
    "FileFormat",
    "PackageValidationError",
    "ValidationError",
    "ValidationErrorType",
    "ValidationResult",
    "ValidationSeverity",
    "ZipPackage",
]
