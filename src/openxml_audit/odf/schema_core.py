"""Schema-core routing and resolver helpers for ODF Relax NG validation."""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass
from fnmatch import fnmatch
from pathlib import Path
from urllib.parse import urlparse

from lxml import etree

from openxml_audit.errors import FileFormat
from openxml_audit.odf._helpers import OFFICE_NS
from openxml_audit.odf.package import OdfPackage

RNG_NS = "http://relaxng.org/ns/structure/1.0"
_RNG_REF_TAGS = {
    f"{{{RNG_NS}}}include",
    f"{{{RNG_NS}}}externalRef",
}
_VERSION_RE = re.compile(r"^\s*(\d+\.\d+)")


def normalize_version_marker(value: str | None) -> str | None:
    """Normalize ODF version markers like '1.3' or '1.4+csd01' to major.minor."""
    if value is None:
        return None
    match = _VERSION_RE.match(value)
    if match is None:
        return None
    return match.group(1)


def default_version_for_format(file_format: FileFormat) -> str:
    """Return the schema-routing version marker for the selected file format."""
    if file_format == FileFormat.ODF_1_2:
        return "1.2"
    return "1.3"


def detect_odf_schema_version(
    package: OdfPackage,
    parsed_parts: Mapping[str, etree._Element],
    *,
    fallback_format: FileFormat,
) -> str:
    """Detect ODF schema version marker from package metadata and XML roots."""
    for part in ("content.xml", "styles.xml", "meta.xml", "settings.xml"):
        root = parsed_parts.get(part)
        if root is None:
            continue
        normalized = normalize_version_marker(root.get(f"{{{OFFICE_NS}}}version"))
        if normalized is not None:
            return normalized

    manifest_version = normalize_version_marker(package.manifest_version)
    if manifest_version is not None:
        return manifest_version
    return default_version_for_format(fallback_format)


@dataclass(frozen=True)
class OdfSchemaRoute:
    """Resolved schema route for a specific part + version."""

    version: str
    pattern: str
    schema_path: Path


class OdfRelaxNgRouter:
    """Resolve Relax NG schema paths for ODF package members."""

    def __init__(self, routes: Mapping[str, Mapping[str, str | Path]] | None = None):
        normalized: dict[str, dict[str, Path]] = {}
        for version, mapping in (routes or {}).items():
            version_key = version.strip() or "*"
            normalized_map: dict[str, Path] = {}
            for pattern, schema in mapping.items():
                pattern_key = pattern.strip()
                if not pattern_key:
                    continue
                normalized_map[pattern_key] = Path(schema)
            if normalized_map:
                normalized[version_key] = normalized_map
        self._routes = normalized

    @classmethod
    def from_legacy_mapping(
        cls,
        mapping: Mapping[str, str | Path] | None,
    ) -> OdfRelaxNgRouter:
        if not mapping:
            return cls()
        return cls({"*": dict(mapping)})

    @classmethod
    def from_bundled(cls) -> OdfRelaxNgRouter:
        """Create a router using bundled structural Relax NG schemas."""
        from openxml_audit.odf.schemas import build_bundled_routes

        return cls(build_bundled_routes())

    def is_empty(self) -> bool:
        return not self._routes

    def resolve(self, part: str, version: str) -> OdfSchemaRoute | None:
        """Resolve a schema route for a package part and version marker."""
        normalized_part = part.lstrip("/")
        prefixed_part = f"/{normalized_part}"
        for version_key in self._candidate_versions(version):
            mapping = self._routes.get(version_key)
            if mapping is None:
                continue

            direct = mapping.get(normalized_part)
            if direct is None:
                direct = mapping.get(prefixed_part)
            if direct is not None:
                return OdfSchemaRoute(
                    version=version_key,
                    pattern=normalized_part,
                    schema_path=direct,
                )

            best_match_pattern: str | None = None
            best_match_schema: Path | None = None
            best_match_len = -1
            for pattern, schema_path in mapping.items():
                if (fnmatch(normalized_part, pattern) or fnmatch(prefixed_part, pattern)) and len(
                    pattern
                ) > best_match_len:
                    best_match_pattern = pattern
                    best_match_schema = schema_path
                    best_match_len = len(pattern)

            if best_match_pattern is not None and best_match_schema is not None:
                return OdfSchemaRoute(
                    version=version_key,
                    pattern=best_match_pattern,
                    schema_path=best_match_schema,
                )
        return None

    @staticmethod
    def _candidate_versions(version: str) -> tuple[str, ...]:
        normalized = normalize_version_marker(version) or version.strip() or "*"
        candidates: list[str] = [normalized]
        if normalized == "1.4":
            candidates.append("1.3")
        candidates.append("*")
        deduped: list[str] = []
        seen: set[str] = set()
        for candidate in candidates:
            if candidate in seen:
                continue
            seen.add(candidate)
            deduped.append(candidate)
        return tuple(deduped)


class OdfRelaxNgResolver:
    """Preflight Relax NG include/externalRef references for deterministic diagnostics."""

    def __init__(self) -> None:
        self._tree_cache: dict[Path, etree._ElementTree] = {}
        self._reference_cache: dict[Path, tuple[str, ...]] = {}
        self._parser = etree.XMLParser(no_network=True, resolve_entities=False)

    def preflight_references(self, schema_path: Path) -> list[str]:
        """Validate include/externalRef references recursively."""
        root = schema_path.resolve()
        cached = self._reference_cache.get(root)
        if cached is not None:
            return list(cached)

        errors: list[str] = []
        self._walk_references(
            schema_path=root,
            ancestry=(),
            visiting=set(),
            visited=set(),
            errors=errors,
        )
        deduped = tuple(dict.fromkeys(errors))
        self._reference_cache[root] = deduped
        return list(deduped)

    def parse_schema(self, schema_path: Path) -> etree._ElementTree:
        """Parse and cache a Relax NG schema tree."""
        normalized = schema_path.resolve()
        cached = self._tree_cache.get(normalized)
        if cached is not None:
            return cached

        tree = etree.parse(str(normalized), parser=self._parser)
        self._tree_cache[normalized] = tree
        return tree

    def _walk_references(
        self,
        *,
        schema_path: Path,
        ancestry: tuple[Path, ...],
        visiting: set[Path],
        visited: set[Path],
        errors: list[str],
    ) -> None:
        if schema_path in visiting:
            cycle = " -> ".join(str(path) for path in (*ancestry, schema_path))
            errors.append(f"Circular Relax NG reference detected: {cycle}")
            return
        if schema_path in visited:
            return

        try:
            tree = self.parse_schema(schema_path)
        except (etree.XMLSyntaxError, OSError) as exc:
            errors.append(f"Invalid Relax NG schema '{schema_path}': {exc}")
            return

        visiting.add(schema_path)
        root = tree.getroot()
        for candidate in root.iter():
            if candidate.tag not in _RNG_REF_TAGS:
                continue
            href = (candidate.get("href") or "").strip()
            if not href:
                errors.append(f"Relax NG reference in '{schema_path}' has empty href attribute")
                continue

            target = self._resolve_reference(schema_path.parent, href)
            if target is None:
                errors.append(f"Unsupported Relax NG reference URI '{href}' in '{schema_path}'")
                continue
            if not target.exists():
                errors.append(f"Unresolvable Relax NG reference '{href}' from '{schema_path}'")
                continue
            self._walk_references(
                schema_path=target,
                ancestry=(*ancestry, schema_path),
                visiting=visiting,
                visited=visited,
                errors=errors,
            )

        visiting.discard(schema_path)
        visited.add(schema_path)

    @staticmethod
    def _resolve_reference(base_dir: Path, href: str) -> Path | None:
        parsed = urlparse(href)
        if parsed.scheme and parsed.scheme != "file":
            return None
        if parsed.scheme == "file":
            return Path(parsed.path).resolve()
        return (base_dir / href).resolve()
