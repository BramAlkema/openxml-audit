"""Euro-Office format capabilities and Document Server client."""

from openxml_audit.eurooffice.client import (
    ConversionResult,
    EuroOfficeClient,
    EuroOfficeConfigError,
    EuroOfficeConversionError,
    EuroOfficeError,
    EuroOfficeRequestError,
)
from openxml_audit.eurooffice.compatibility import (
    TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
    EuroOfficeCompatibilityReport,
    EuroOfficeCompatibilityStatus,
    classify_eurooffice_compatibility,
    supported_eurooffice_compatibility_profiles,
)
from openxml_audit.eurooffice.formats import (
    CONNECTOR_VERSION,
    DOCUMENT_FORMATS_COMMIT,
    DOCUMENT_FORMATS_VERSION,
    DOCUMENT_SERVER_RELEASE,
    EuroOfficeFormatMode,
    EuroOfficeFormatSupport,
    format_support,
)

__all__ = [
    "CONNECTOR_VERSION",
    "DOCUMENT_SERVER_RELEASE",
    "DOCUMENT_FORMATS_COMMIT",
    "DOCUMENT_FORMATS_VERSION",
    "TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID",
    "ConversionResult",
    "EuroOfficeClient",
    "EuroOfficeCompatibilityReport",
    "EuroOfficeCompatibilityStatus",
    "EuroOfficeConfigError",
    "EuroOfficeConversionError",
    "EuroOfficeError",
    "EuroOfficeFormatMode",
    "EuroOfficeFormatSupport",
    "EuroOfficeRequestError",
    "classify_eurooffice_compatibility",
    "format_support",
    "supported_eurooffice_compatibility_profiles",
]
