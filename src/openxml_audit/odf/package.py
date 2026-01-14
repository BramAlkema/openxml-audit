"""ODF package handling."""

from __future__ import annotations

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

            ns = {"manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"}
            entries: list[OdfManifestEntry] = []
            for entry in xml.findall("manifest:file-entry", ns):
                full_path = entry.get(f"{{{ns['manifest']}}}full-path", "")
                media_type = entry.get(f"{{{ns['manifest']}}}media-type", "")
                entries.append(OdfManifestEntry(full_path=full_path, media_type=media_type))

            self._manifest = entries
        return self._manifest

    def list_xml_parts(self) -> Iterator[str]:
        """List manifest parts that look like XML."""
        for entry in self.manifest:
            if entry.media_type.endswith("xml"):
                yield entry.full_path

    def validate_structure(self) -> list[ValidationError]:
        """Perform basic ODF package checks."""
        errors: list[ValidationError] = []

        _ = self.mimetype
        _ = self.manifest
        errors.extend(self._errors)

        return errors
