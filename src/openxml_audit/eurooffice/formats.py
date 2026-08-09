"""Euro-Office/ONLYOFFICE connector format capability matrix.

The matrix is pinned to Nextcloud connector 11.0.1 and its
ONLYOFFICE/document-formats 3.2.0 dependency.  It deliberately distinguishes
native OOXML editing from ODF editing through lossy conversion: an ``.odt``
that appears editable in Nextcloud is converted to ``.docx`` before editing.
``.odg`` is available for viewing/conversion, but is not an editable format.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from urllib.parse import urlparse

CONNECTOR_VERSION = "11.0.1"
DOCUMENT_SERVER_RELEASE = "9.3.3"
DOCUMENT_FORMATS_VERSION = "3.2.0"
DOCUMENT_FORMATS_COMMIT = "7d7576a3fe2337c30f4c9b40fae70a69dc68ba08"


class EuroOfficeFormatMode(str, Enum):
    """How the current connector exposes a format."""

    NATIVE_EDIT = "native-edit"
    LOSSY_EDIT = "lossy-edit"
    VIEW_ONLY = "view-only"
    UNSUPPORTED = "unsupported"


@dataclass(frozen=True)
class EuroOfficeFormatSupport:
    """Connector capability for one filename extension."""

    extension: str
    document_type: str
    mode: EuroOfficeFormatMode
    conversion_target: str | None
    actions: tuple[str, ...]


_NATIVE_EDIT: dict[str, str] = {
    "docx": "word",
    "docm": "word",
    "dotx": "word",
    "dotm": "word",
    "xlsx": "cell",
    "xlsm": "cell",
    "xltx": "cell",
    "xltm": "cell",
    "pptx": "slide",
    "pptm": "slide",
    "potx": "slide",
    "potm": "slide",
    "ppsx": "slide",
    "ppsm": "slide",
}

_LOSSY_EDIT: dict[str, tuple[str, str]] = {
    "odt": ("word", "docx"),
    "ott": ("word", "docx"),
    "ods": ("cell", "xlsx"),
    "ots": ("cell", "xlsx"),
    "odp": ("slide", "pptx"),
    "otp": ("slide", "pptx"),
}

_VIEW_ONLY: dict[str, tuple[str, str]] = {
    "odg": ("slide", "pptx"),
}


def _extension(value: str | Path) -> str:
    raw = str(value).strip()
    if not raw:
        return ""
    if raw.startswith(".") and "/" not in raw and "\\" not in raw:
        return raw[1:].lower()

    parsed_path = urlparse(raw).path if "://" in raw else raw
    suffix = Path(parsed_path).suffix
    if suffix:
        return suffix[1:].lower()
    if "/" not in raw and "\\" not in raw:
        return raw.lower()
    return ""


def format_support(value: str | Path) -> EuroOfficeFormatSupport:
    """Return the pinned connector capability for a path or extension."""

    extension = _extension(value)
    if extension in _NATIVE_EDIT:
        return EuroOfficeFormatSupport(
            extension=extension,
            document_type=_NATIVE_EDIT[extension],
            mode=EuroOfficeFormatMode.NATIVE_EDIT,
            conversion_target=extension,
            actions=("view", "edit"),
        )
    if extension in _LOSSY_EDIT:
        document_type, target = _LOSSY_EDIT[extension]
        return EuroOfficeFormatSupport(
            extension=extension,
            document_type=document_type,
            mode=EuroOfficeFormatMode.LOSSY_EDIT,
            conversion_target=target,
            actions=("view", "edit", "auto-convert"),
        )
    if extension in _VIEW_ONLY:
        document_type, target = _VIEW_ONLY[extension]
        return EuroOfficeFormatSupport(
            extension=extension,
            document_type=document_type,
            mode=EuroOfficeFormatMode.VIEW_ONLY,
            conversion_target=target,
            actions=("view", "auto-convert"),
        )
    return EuroOfficeFormatSupport(
        extension=extension,
        document_type="unknown",
        mode=EuroOfficeFormatMode.UNSUPPORTED,
        conversion_target=None,
        actions=(),
    )


__all__ = [
    "CONNECTOR_VERSION",
    "DOCUMENT_SERVER_RELEASE",
    "DOCUMENT_FORMATS_COMMIT",
    "DOCUMENT_FORMATS_VERSION",
    "EuroOfficeFormatMode",
    "EuroOfficeFormatSupport",
    "format_support",
]
