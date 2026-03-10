"""Bundled OASIS ODF Relax NG schemas for zero-config validation.

Ships the official OASIS Relax NG schemas for ODF 1.2 and 1.3.
These are the full schemas — not subsets — and are redistributed under the
OASIS copyright notice included in each schema file.

Schema files per version:
  - ``OpenDocument-v{VER}-schema.rng``          — content, styles, meta, settings
  - ``OpenDocument-v{VER}-manifest-schema.rng``  — META-INF/manifest.xml
  - ``OpenDocument-v{VER}-dsig-schema.rng``      — digital signatures
"""

from __future__ import annotations

from pathlib import Path

_SCHEMAS_DIR = Path(__file__).parent

# Parts validated by the main document schema
_DOCUMENT_PARTS = ("content.xml", "styles.xml", "meta.xml", "settings.xml")
_MANIFEST_PART = "META-INF/manifest.xml"

# Supported bundled versions
BUNDLED_VERSIONS = ("1.2", "1.3")

# Schema filename templates per version
_SCHEMA_FILENAMES: dict[str, dict[str, str]] = {
    "1.2": {
        "document": "OpenDocument-v1.2-os-schema.rng",
        "manifest": "OpenDocument-v1.2-os-manifest-schema.rng",
        "dsig": "OpenDocument-v1.2-os-dsig-schema.rng",
    },
    "1.3": {
        "document": "OpenDocument-v1.3-schema.rng",
        "manifest": "OpenDocument-v1.3-manifest-schema.rng",
        "dsig": "OpenDocument-v1.3-dsig-schema.rng",
    },
}


def get_bundled_schema_dir(version: str) -> Path | None:
    """Return the directory containing bundled schemas for a version."""
    if version in BUNDLED_VERSIONS:
        return _SCHEMAS_DIR / f"odf-{version}"
    return None


def get_bundled_schema_path(part: str, version: str) -> Path | None:
    """Return the path to a bundled schema for a specific part and version."""
    schema_dir = get_bundled_schema_dir(version)
    if schema_dir is None or version not in _SCHEMA_FILENAMES:
        return None

    filenames = _SCHEMA_FILENAMES[version]
    normalized = part.lstrip("/")

    if normalized in _DOCUMENT_PARTS:
        candidate = schema_dir / filenames["document"]
    elif normalized == _MANIFEST_PART:
        candidate = schema_dir / filenames["manifest"]
    else:
        return None

    return candidate if candidate.exists() else None


def build_bundled_routes() -> dict[str, dict[str, Path]]:
    """Build schema routes mapping for all bundled versions.

    Returns a dict suitable for ``OdfRelaxNgRouter(routes)``.
    The main OASIS schema validates content, styles, meta, and settings.
    A separate manifest schema validates META-INF/manifest.xml.
    """
    routes: dict[str, dict[str, Path]] = {}
    for version in BUNDLED_VERSIONS:
        schema_dir = get_bundled_schema_dir(version)
        if schema_dir is None or version not in _SCHEMA_FILENAMES:
            continue
        filenames = _SCHEMA_FILENAMES[version]
        version_map: dict[str, Path] = {}

        # Main document schema — used for all document parts
        doc_schema = schema_dir / filenames["document"]
        if doc_schema.exists():
            for part in _DOCUMENT_PARTS:
                version_map[part] = doc_schema

        # Manifest schema
        manifest_schema = schema_dir / filenames["manifest"]
        if manifest_schema.exists():
            version_map[_MANIFEST_PART] = manifest_schema

        if version_map:
            routes[version] = version_map
    return routes
