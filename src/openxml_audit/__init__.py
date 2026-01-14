"""OpenXML Audit - Python port of Open XML SDK validation.

Validate if OOXML files will open in Microsoft Office.

Example:
    from openxml_audit import validate_pptx, is_valid_pptx

    # Quick check
    if is_valid_pptx("presentation.pptx"):
        print("File is valid!")

    # Detailed validation
    result = validate_pptx("presentation.pptx")
    if not result.is_valid:
        for error in result.errors:
            print(error)

    # With custom options
    from openxml_audit import OpenXmlValidator, FileFormat

    validator = OpenXmlValidator(
        file_format=FileFormat.OFFICE_2019,
        max_errors=100,
    )
    result = validator.validate("presentation.pptx")
"""

from openxml_audit.errors import (
    FileFormat,
    PackageValidationError,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.package import OpenXmlPackage
from openxml_audit.odf import OdfPackage, OdfValidator
from openxml_audit.parts import (
    DocumentPart,
    OpenXmlPart,
    PresentationPart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
    ThemePart,
    WorkbookPart,
)
from openxml_audit.validator import OpenXmlValidator, is_valid_pptx, validate_pptx
from openxml_audit.helpers import (
    require_valid_pptx,
    validate_on_save,
    validation_context,
)

__version__ = "0.1.0"

__all__ = [
    # Main API
    "OpenXmlValidator",
    "validate_pptx",
    "is_valid_pptx",
    # Results and errors
    "ValidationResult",
    "ValidationError",
    "ValidationErrorType",
    "ValidationSeverity",
    "PackageValidationError",
    # Enums
    "FileFormat",
    # Package and parts (for advanced usage)
    "OpenXmlPackage",
    "OdfPackage",
    "OpenXmlPart",
    "DocumentPart",
    "PresentationPart",
    "SlidePart",
    "SlideLayoutPart",
    "SlideMasterPart",
    "ThemePart",
    "WorkbookPart",
    # ODF validator
    "OdfValidator",
    # Integration helpers
    "validation_context",
    "validate_on_save",
    "require_valid_pptx",
]
