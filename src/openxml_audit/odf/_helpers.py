"""Shared ODF helpers used across validator, semantic, and security modules."""

from __future__ import annotations

# ODF namespace constants
OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
XMLDSIG_NS = "http://www.w3.org/2000/09/xmldsig#"
SIGNATURE_PKG_NS = "urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"
TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
META_NS = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
NUMBER_NS = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
XLINK_NS = "http://www.w3.org/1999/xlink"


def normalize_part_uri(part_path: str) -> str:
    """Ensure part URI has a leading slash."""
    return part_path if part_path.startswith("/") else f"/{part_path}"


def normalize_manifest_path(path: str) -> str:
    """Normalize a manifest full-path to its canonical form."""
    cleaned = path.strip()
    if not cleaned:
        return ""
    if cleaned == "/":
        return "/"
    return cleaned.lstrip("/")
