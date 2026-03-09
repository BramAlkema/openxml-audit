"""ODF package handling."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.core.package import ZipPackage
from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity

if TYPE_CHECKING:
    from collections.abc import Iterator


@dataclass
class OdfManifestEntry:
    """Entry from META-INF/manifest.xml."""

    full_path: str
    media_type: str


class OdfPackage(ZipPackage):
    """ODF package (ODT/ODS/ODP)."""

    MANIFEST_PATH = "META-INF/manifest.xml"
    MIMETYPE_PATH = "mimetype"
    MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    MIMETYPE_PREFIX = "application/vnd.oasis.opendocument."
    REQUIRED_CONTENT_PREFIXES = (
        "application/vnd.oasis.opendocument.text",
        "application/vnd.oasis.opendocument.spreadsheet",
        "application/vnd.oasis.opendocument.presentation",
    )

    def __init__(self, path: str | Path):
        super().__init__(path)
        self._manifest: list[OdfManifestEntry] | None = None
        self._mimetype: str | None = None

    @property
    def mimetype(self) -> str | None:
        """Get the mimetype declared by the package."""
        if self._mimetype is None:
            content = self.get_part_content(self.MIMETYPE_PATH)
            if content is None:
                self._errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description="Missing mimetype entry",
                        severity=ValidationSeverity.ERROR,
                    )
                )
                return None
            self._mimetype = content.decode("utf-8", errors="replace").strip()
        return self._mimetype

    @property
    def manifest(self) -> list[OdfManifestEntry]:
        """Get the parsed manifest entries."""
        if self._manifest is None:
            content = self.get_part_content(self.MANIFEST_PATH)
            if content is None:
                self._errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description="Missing META-INF/manifest.xml",
                        severity=ValidationSeverity.ERROR,
                    )
                )
                self._manifest = []
                return self._manifest

            try:
                xml = etree.fromstring(content)
            except etree.XMLSyntaxError as exc:
                self._errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.SCHEMA,
                        description=f"Invalid manifest.xml: {exc}",
                        part_uri=self.MANIFEST_PATH,
                        severity=ValidationSeverity.ERROR,
                    )
                )
                self._manifest = []
                return self._manifest

            ns = {"manifest": self.MANIFEST_NS}
            entries: list[OdfManifestEntry] = []
            for entry in xml.findall("manifest:file-entry", ns):
                full_path = entry.get(f"{{{ns['manifest']}}}full-path", "")
                media_type = entry.get(f"{{{ns['manifest']}}}media-type", "")
                entries.append(OdfManifestEntry(full_path=full_path, media_type=media_type))

            self._manifest = entries
        return self._manifest

    @staticmethod
    def _normalize_manifest_path(path: str) -> str:
        cleaned = path.strip()
        if not cleaned:
            return ""
        if cleaned == "/":
            return "/"
        return cleaned.lstrip("/")

    def _zip_members(self) -> set[str]:
        return {part.lstrip("/") for part in super().list_parts()}

    @staticmethod
    def _is_xml_manifest_entry(entry: OdfManifestEntry) -> bool:
        media_type = entry.media_type.strip().lower()
        path = entry.full_path.strip().lower()
        return media_type.endswith("xml") or path.endswith(".xml")

    @staticmethod
    def _conformance_severity(strict: bool) -> ValidationSeverity:
        return ValidationSeverity.ERROR if strict else ValidationSeverity.WARNING

    @staticmethod
    def _is_supported_mimetype(mimetype: str) -> bool:
        return mimetype.startswith(OdfPackage.MIMETYPE_PREFIX)

    @staticmethod
    def _requires_content_xml(mimetype: str) -> bool:
        return mimetype.startswith(OdfPackage.REQUIRED_CONTENT_PREFIXES)

    def list_xml_parts(self) -> Iterator[str]:
        """List manifest parts that look like XML."""
        seen: set[str] = set()
        for entry in self.manifest:
            path = self._normalize_manifest_path(entry.full_path)
            if path in {"", "/"} or path.endswith("/"):
                continue
            if not self._is_xml_manifest_entry(entry):
                continue
            if path in seen:
                continue
            seen.add(path)
            yield path

    def manifest_paths(self) -> set[str]:
        """Return normalized non-root manifest member paths."""
        return {
            self._normalize_manifest_path(entry.full_path)
            for entry in self.manifest
            if self._normalize_manifest_path(entry.full_path) not in {"", "/"}
        }

    def validate_structure(self, strict: bool = True) -> list[ValidationError]:
        """Perform basic ODF package checks."""
        _ = self.mimetype
        _ = self.manifest
        errors: list[ValidationError] = list(self._errors)

        mimetype = self._mimetype
        manifest = self._manifest or []

        if mimetype is None or not manifest:
            return errors

        if not self._is_supported_mimetype(mimetype):
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=f"Invalid ODF mimetype value '{mimetype}'",
                    part_uri=self.MIMETYPE_PATH,
                    severity=ValidationSeverity.ERROR,
                )
            )

        normalized_paths: list[str] = [
            self._normalize_manifest_path(entry.full_path) for entry in manifest
        ]
        counts = Counter(path for path in normalized_paths if path)
        for path, count in sorted(counts.items()):
            if count <= 1:
                continue
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=f"Duplicate manifest file-entry path '{path}' ({count} entries)",
                    part_uri=self.MANIFEST_PATH,
                    severity=self._conformance_severity(strict),
                )
            )

        root_entries = [
            entry
            for entry in manifest
            if self._normalize_manifest_path(entry.full_path) == "/"
        ]
        if not root_entries:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description="manifest.xml is missing required root file-entry '/'",
                    part_uri=self.MANIFEST_PATH,
                    severity=self._conformance_severity(strict),
                )
            )
        elif mimetype is not None:
            root_media_type = root_entries[0].media_type.strip()
            if root_media_type != mimetype:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=(
                            "Root manifest media-type does not match mimetype "
                            f"('{root_media_type}' != '{mimetype}')"
                        ),
                        part_uri=self.MANIFEST_PATH,
                        severity=self._conformance_severity(strict),
                    )
                )

        zip_members = self._zip_members()
        manifest_members = self.manifest_paths()

        for member in sorted(manifest_members):
            if member.endswith("/"):
                continue
            if member not in zip_members:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=f"Manifest entry '{member}' was not found in package",
                        part_uri=self.MANIFEST_PATH,
                        severity=self._conformance_severity(strict),
                    )
                )

        for member in sorted(zip_members):
            if member in {self.MIMETYPE_PATH, self.MANIFEST_PATH}:
                continue
            if member.startswith("META-INF/"):
                continue
            if member not in manifest_members:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=f"Package member '{member}' is not declared in manifest.xml",
                        part_uri=self.MANIFEST_PATH,
                        severity=self._conformance_severity(strict),
                    )
                )

        if self._requires_content_xml(mimetype):
            if "content.xml" not in zip_members:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description="ODF package is missing required member 'content.xml'",
                        part_uri=self.MANIFEST_PATH,
                        severity=self._conformance_severity(strict),
                    )
                )
            if "content.xml" not in manifest_members:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description="ODF package is missing required manifest entry 'content.xml'",
                        part_uri=self.MANIFEST_PATH,
                        severity=self._conformance_severity(strict),
                    )
                )

        return errors
