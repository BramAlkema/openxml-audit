"""Binary payload validation for embedded parts."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from openxml_audit.errors import ValidationSeverity


@dataclass(frozen=True)
class BinaryFormat:
    """Definition of a binary payload format."""

    name: str
    content_types: tuple[str, ...]
    extensions: tuple[str, ...]
    validator: Callable[[bytes], bool]


@dataclass(frozen=True)
class BinaryValidationResult:
    """Validation outcome for a binary payload."""

    message: str
    severity: ValidationSeverity = ValidationSeverity.ERROR


JPEG_MAGIC = (b"\xFF\xD8\xFF",)
PNG_MAGIC = (b"\x89PNG\r\n\x1a\n",)
GIF_MAGIC = (b"GIF87a", b"GIF89a")
BMP_MAGIC = (b"BM",)
TIFF_MAGIC = (b"II*\x00", b"MM\x00*")
OLE_MAGIC = (b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1",)
WMF_PLACEABLE_MAGIC = b"\xD7\xCD\xC6\x9A"
FONT_MAGIC = (b"\x00\x01\x00\x00", b"OTTO", b"ttcf", b"true", b"typ1")

FONT_CONTENT_TYPES = (
    "application/vnd.ms-opentype",
    "application/x-font-ttf",
    "application/x-font-opentype",
    "application/x-fontdata",
)
OBFUSCATED_FONT_CONTENT_TYPES = (
    "application/vnd.openxmlformats-officedocument.obfuscatedFont",
)
FONT_EXTENSIONS = (".ttf", ".otf", ".ttc", ".otc", ".fntdata", ".odttf")
OBFUSCATED_FONT_EXTENSIONS = (".odttf",)


def _starts_with_any(data: bytes, candidates: tuple[bytes, ...]) -> bool:
    return any(data.startswith(prefix) for prefix in candidates)


def _is_jpeg(data: bytes) -> bool:
    return _starts_with_any(data, JPEG_MAGIC)


def _is_png(data: bytes) -> bool:
    return _starts_with_any(data, PNG_MAGIC)


def _is_gif(data: bytes) -> bool:
    return _starts_with_any(data, GIF_MAGIC)


def _is_bmp(data: bytes) -> bool:
    return _starts_with_any(data, BMP_MAGIC)


def _is_tiff(data: bytes) -> bool:
    return _starts_with_any(data, TIFF_MAGIC)


def _is_emf(data: bytes) -> bool:
    if len(data) < 44:
        return False
    # EMF signature appears at offset 40 as " EMF".
    return data[:4] == b"\x01\x00\x00\x00" and data[40:44] == b" EMF"


def _is_wmf(data: bytes) -> bool:
    if len(data) < 4:
        return False
    if data.startswith(WMF_PLACEABLE_MAGIC):
        return True
    # Non-placeable WMF header: type (1 or 2) + header size (9)
    return data[:2] in (b"\x01\x00", b"\x02\x00") and data[2:4] == b"\x09\x00"


def _is_ole(data: bytes) -> bool:
    return _starts_with_any(data, OLE_MAGIC)


def _is_font_header(data: bytes) -> bool:
    return _starts_with_any(data, FONT_MAGIC)


def _extract_fntdata_payload(data: bytes) -> bytes | None:
    if len(data) < 8:
        return None
    total = int.from_bytes(data[0:4], "little")
    font_len = int.from_bytes(data[4:8], "little")
    if total <= 0 or font_len <= 0:
        return None
    if total > len(data):
        return None
    offset = total - font_len
    if offset < 8 or offset >= len(data):
        return None
    return data[offset:]


def _get_extension(part_uri: str) -> str:
    ext = part_uri.lower().rsplit(".", 1)[-1] if "." in part_uri else ""
    return f".{ext}" if ext else ""


def _is_font_candidate(content_type: str | None, part_uri: str) -> bool:
    ext = _get_extension(part_uri)
    if content_type:
        if content_type in FONT_CONTENT_TYPES or content_type in OBFUSCATED_FONT_CONTENT_TYPES:
            return True
    return ext in FONT_EXTENSIONS


def _is_obfuscated_font(content_type: str | None, part_uri: str) -> bool:
    ext = _get_extension(part_uri)
    if content_type and content_type in OBFUSCATED_FONT_CONTENT_TYPES:
        return True
    return ext in OBFUSCATED_FONT_EXTENSIONS


def parse_font_key(value: str) -> bytes | None:
    """Parse a GUID-style font key into bytes for deobfuscation."""
    text = value.strip().lower()
    if text.startswith("{") and text.endswith("}"):
        text = text[1:-1]
    parts = text.split("-")
    if len(parts) != 5:
        return None
    lengths = (8, 4, 4, 4, 12)
    if any(len(part) != expected for part, expected in zip(parts, lengths)):
        return None
    try:
        data1 = bytes.fromhex(parts[0])
        data2 = bytes.fromhex(parts[1])
        data3 = bytes.fromhex(parts[2])
        data4 = bytes.fromhex(parts[3] + parts[4])
    except ValueError:
        return None
    if not (len(data1) == 4 and len(data2) == 2 and len(data3) == 2 and len(data4) == 8):
        return None
    return data1[::-1] + data2[::-1] + data3[::-1] + data4


def _deobfuscate_prefix(data: bytes, key: bytes, length: int = 32) -> bytes:
    limit = min(len(data), length)
    if limit == 0:
        return b""
    if len(key) != 16:
        return data[:limit]
    return bytes(data[i] ^ key[i % 16] for i in range(limit))


BINARY_FORMATS: tuple[BinaryFormat, ...] = (
    BinaryFormat(
        name="jpeg",
        content_types=("image/jpeg", "image/pjpeg"),
        extensions=(".jpg", ".jpeg"),
        validator=_is_jpeg,
    ),
    BinaryFormat(
        name="png",
        content_types=("image/png",),
        extensions=(".png",),
        validator=_is_png,
    ),
    BinaryFormat(
        name="gif",
        content_types=("image/gif",),
        extensions=(".gif",),
        validator=_is_gif,
    ),
    BinaryFormat(
        name="bmp",
        content_types=("image/bmp", "image/x-bmp"),
        extensions=(".bmp",),
        validator=_is_bmp,
    ),
    BinaryFormat(
        name="tiff",
        content_types=("image/tiff",),
        extensions=(".tif", ".tiff"),
        validator=_is_tiff,
    ),
    BinaryFormat(
        name="emf",
        content_types=("image/emf", "image/x-emf"),
        extensions=(".emf",),
        validator=_is_emf,
    ),
    BinaryFormat(
        name="wmf",
        content_types=("image/wmf", "image/x-wmf"),
        extensions=(".wmf",),
        validator=_is_wmf,
    ),
    BinaryFormat(
        name="ole",
        content_types=(
            "application/vnd.openxmlformats-officedocument.oleObject",
            "application/vnd.ms-office.activeX",
        ),
        extensions=(".bin", ".ole"),
        validator=_is_ole,
    ),
)


def _match_format(content_type: str | None, part_uri: str) -> BinaryFormat | None:
    ext = _get_extension(part_uri)
    for fmt in BINARY_FORMATS:
        if content_type and content_type in fmt.content_types:
            return fmt
        if ext and ext in fmt.extensions:
            return fmt
    return None


def validate_binary_content(
    content_type: str | None,
    part_uri: str,
    data: bytes,
    font_key: bytes | None = None,
) -> BinaryValidationResult | None:
    """Validate binary payload based on content type or extension.

    Returns a validation result if invalid, otherwise None.
    """
    if _is_font_candidate(content_type, part_uri):
        if _get_extension(part_uri) == ".fntdata" or content_type == "application/x-fontdata":
            payload = _extract_fntdata_payload(data)
            if payload and _is_font_header(payload):
                return None
        if _is_font_header(data):
            return None
        if _is_obfuscated_font(content_type, part_uri):
            if font_key is None or len(font_key) != 16:
                return BinaryValidationResult(
                    "Obfuscated font payload missing fontKey; unable to validate.",
                    severity=ValidationSeverity.WARNING,
                )
            deobfuscated = _deobfuscate_prefix(data, font_key)
            if _is_font_header(deobfuscated):
                return None
            return BinaryValidationResult("Invalid obfuscated font payload after deobfuscation.")
        return BinaryValidationResult("Invalid font payload.")

    fmt = _match_format(content_type, part_uri)
    if fmt is None:
        return None
    if fmt.validator(data):
        return None
    ct_hint = f" (content type {content_type})" if content_type else ""
    return BinaryValidationResult(f"Invalid {fmt.name} payload{ct_hint}.")
