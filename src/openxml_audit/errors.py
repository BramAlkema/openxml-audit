"""Validation error types and file format definitions."""

from dataclasses import dataclass, field
from enum import Enum


class FileFormat(Enum):
    """Office file format versions for validation."""

    OFFICE_2007 = "office2007"  # ECMA-376 1st edition
    OFFICE_2010 = "office2010"  # ECMA-376 2nd edition
    OFFICE_2013 = "office2013"
    OFFICE_2016 = "office2016"
    OFFICE_2019 = "office2019"
    OFFICE_2021 = "office2021"
    MICROSOFT_365 = "microsoft365"
    ODF_1_2 = "odf1.2"
    ODF_1_3 = "odf1.3"

    @classmethod
    def from_version_string(cls, version: str) -> "FileFormat | None":
        """Map SDK version strings like 'Office2010' to FileFormat."""
        return _VERSION_STRING_MAP.get(version)

    def ooxml_order(self) -> int:
        """Return ordinal position for OOXML version comparison.

        Returns -1 for non-OOXML formats.
        """
        return _OOXML_ORDER.get(self, -1)

    def includes_ooxml(self, other: "FileFormat") -> bool:
        """Check if this format version includes elements from `other`.

        Returns True if `other` was introduced at or before this version.
        """
        self_order = self.ooxml_order()
        other_order = other.ooxml_order()
        if self_order < 0 or other_order < 0:
            return False
        return self_order >= other_order


_VERSION_STRING_MAP: dict[str, "FileFormat"] = {
    "Office2007": FileFormat.OFFICE_2007,
    "Office2010": FileFormat.OFFICE_2010,
    "Office2013": FileFormat.OFFICE_2013,
    "Office2016": FileFormat.OFFICE_2016,
    "Office2019": FileFormat.OFFICE_2019,
    "Office2021": FileFormat.OFFICE_2021,
    "Microsoft365": FileFormat.MICROSOFT_365,
}

_OOXML_ORDER: dict["FileFormat", int] = {
    FileFormat.OFFICE_2007: 0,
    FileFormat.OFFICE_2010: 1,
    FileFormat.OFFICE_2013: 2,
    FileFormat.OFFICE_2016: 3,
    FileFormat.OFFICE_2019: 4,
    FileFormat.OFFICE_2021: 5,
    FileFormat.MICROSOFT_365: 6,
}


class ValidationErrorType(Enum):
    """Types of validation errors."""

    PACKAGE = "package"  # OPC package structure error
    BINARY = "binary"  # Binary payload validation error
    SCHEMA = "schema"  # XML schema violation
    SEMANTIC = "semantic"  # Semantic constraint violation
    RELATIONSHIP = "relationship"  # Relationship error
    MARKUP_COMPATIBILITY = "markup_compatibility"  # MC error


class ValidationSeverity(Enum):
    """Severity levels for validation errors."""

    ERROR = "error"  # Will prevent file from opening
    WARNING = "warning"  # May cause issues
    INFO = "info"  # Informational


@dataclass
class ValidationError:
    """A validation error found in a document."""

    error_type: ValidationErrorType
    description: str
    part_uri: str = ""  # e.g., "/ppt/slides/slide1.xml"
    path: str = ""  # XPath to element
    node: str | None = None  # Element/attribute name
    related_node: str | None = None
    severity: ValidationSeverity = ValidationSeverity.ERROR
    id: str = ""  # Error ID for categorization

    def __str__(self) -> str:
        location = self.part_uri
        if self.path:
            location = f"{location}:{self.path}"
        return f"[{self.error_type.value}] {location}: {self.description}"


@dataclass
class ValidationResult:
    """Result of validating a document."""

    is_valid: bool
    errors: list[ValidationError] = field(default_factory=list)
    file_path: str = ""
    file_format: FileFormat = FileFormat.OFFICE_2019

    @property
    def error_count(self) -> int:
        return sum(1 for e in self.errors if e.severity == ValidationSeverity.ERROR)

    @property
    def warning_count(self) -> int:
        return sum(1 for e in self.errors if e.severity == ValidationSeverity.WARNING)


class PackageValidationError(Exception):
    """Exception raised when package validation fails catastrophically."""

    def __init__(self, message: str, errors: list[ValidationError] | None = None):
        super().__init__(message)
        self.errors = errors or []
