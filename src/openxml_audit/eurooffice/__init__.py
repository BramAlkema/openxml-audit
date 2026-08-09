"""Euro-Office format capabilities and Document Server client."""

from openxml_audit.eurooffice.client import (
    ConversionResult,
    EuroOfficeClient,
    EuroOfficeConfigError,
    EuroOfficeConversionError,
    EuroOfficeError,
    EuroOfficeRequestError,
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
    "ConversionResult",
    "EuroOfficeClient",
    "EuroOfficeConfigError",
    "EuroOfficeConversionError",
    "EuroOfficeError",
    "EuroOfficeFormatMode",
    "EuroOfficeFormatSupport",
    "EuroOfficeRequestError",
    "format_support",
]
