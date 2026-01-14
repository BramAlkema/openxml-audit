"""Generic ZIP package handling shared across validators."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import (
    PackageValidationError,
    ValidationError,
    ValidationErrorType,
    ValidationSeverity,
)

if TYPE_CHECKING:
    from collections.abc import Iterator


class ZipPackage:
    """A ZIP-backed document package."""

    def __init__(self, path: str | Path):
        self._path = Path(path)
        self._zip: zipfile.ZipFile | None = None
        self._part_cache: dict[str, bytes] = {}
        self._errors: list[ValidationError] = []

    def __enter__(self) -> ZipPackage:
        self.open()
        return self

    def __exit__(self, exc_type: object, exc_val: object, exc_tb: object) -> None:
        self.close()

    def open(self) -> None:
        """Open the package for reading."""
        if self._zip is not None:
            return

        try:
            self._zip = zipfile.ZipFile(self._path, "r")
        except zipfile.BadZipFile as exc:
            raise PackageValidationError(
                f"Invalid ZIP file: {exc}",
                [
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=f"File is not a valid ZIP archive: {exc}",
                        severity=ValidationSeverity.ERROR,
                    )
                ],
            ) from exc
        except FileNotFoundError as exc:
            raise PackageValidationError(
                f"File not found: {self._path}",
                [
                    ValidationError(
                        error_type=ValidationErrorType.PACKAGE,
                        description=f"File not found: {self._path}",
                        severity=ValidationSeverity.ERROR,
                    )
                ],
            ) from exc

    def close(self) -> None:
        """Close the package."""
        if self._zip is not None:
            self._zip.close()
            self._zip = None
        self._part_cache.clear()

    @property
    def path(self) -> Path:
        """Get the path to the package file."""
        return self._path

    @property
    def errors(self) -> list[ValidationError]:
        """Get any errors encountered while reading the package."""
        return self._errors

    def get_part_content(self, part_path: str) -> bytes | None:
        """Get the raw content of a part."""
        if self._zip is None:
            raise PackageValidationError("Package not opened")

        zip_path = part_path.lstrip("/")

        if zip_path in self._part_cache:
            return self._part_cache[zip_path]

        try:
            content = self._zip.read(zip_path)
            self._part_cache[zip_path] = content
            return content
        except KeyError:
            return None

    def get_part_xml(self, part_path: str) -> etree._Element | None:
        """Get the parsed XML content of a part."""
        content = self.get_part_content(part_path)
        if content is None:
            return None

        try:
            return etree.fromstring(content)
        except etree.XMLSyntaxError as exc:
            self._errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=f"XML parse error: {exc}",
                    part_uri=part_path,
                    severity=ValidationSeverity.ERROR,
                )
            )
            return None

    def list_parts(self) -> Iterator[str]:
        """List all parts in the package."""
        if self._zip is None:
            raise PackageValidationError("Package not opened")

        for name in self._zip.namelist():
            yield "/" + name

    def has_part(self, part_path: str) -> bool:
        """Check if a part exists in the package."""
        if self._zip is None:
            raise PackageValidationError("Package not opened")

        zip_path = part_path.lstrip("/")
        return zip_path in self._zip.namelist()
