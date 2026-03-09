"""Tests for ODF package structure validation."""

from __future__ import annotations

from pathlib import Path

import pytest

from openxml_audit.errors import ValidationErrorType, ValidationSeverity
from openxml_audit.odf.package import OdfPackage


class TestOdfPackage:
    """ODF package structure checks."""

    @pytest.mark.parametrize("fixture_name", ["minimal_odt", "minimal_ods", "minimal_odp"])
    def test_valid_minimal_packages_pass(
        self,
        request: pytest.FixtureRequest,
        fixture_name: str,
    ) -> None:
        path = request.getfixturevalue(fixture_name)

        with OdfPackage(path) as package:
            errors = package.validate_structure(strict=True)

        assert not any(error.severity == ValidationSeverity.ERROR for error in errors)

    def test_missing_mimetype_reports_error(self, odf_missing_mimetype: Path) -> None:
        with OdfPackage(odf_missing_mimetype) as package:
            errors = package.validate_structure()

        assert any(
            error.error_type == ValidationErrorType.PACKAGE
            and "Missing mimetype entry" in error.description
            for error in errors
        )

    def test_invalid_mimetype_reports_error(self, odf_invalid_mimetype: Path) -> None:
        with OdfPackage(odf_invalid_mimetype) as package:
            errors = package.validate_structure()

        assert any(
            error.error_type == ValidationErrorType.PACKAGE
            and "Invalid ODF mimetype value" in error.description
            for error in errors
        )

    def test_missing_manifest_reports_error(self, odf_missing_manifest: Path) -> None:
        with OdfPackage(odf_missing_manifest) as package:
            errors = package.validate_structure()

        assert any(
            error.error_type == ValidationErrorType.PACKAGE
            and "Missing META-INF/manifest.xml" in error.description
            for error in errors
        )

    def test_malformed_manifest_reports_schema_error(self, odf_malformed_manifest: Path) -> None:
        with OdfPackage(odf_malformed_manifest) as package:
            errors = package.validate_structure()

        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and error.part_uri == "META-INF/manifest.xml"
            and "Invalid manifest.xml" in error.description
            for error in errors
        )

    def test_manifest_reference_to_missing_part_is_reported(
        self, odf_manifest_missing_part: Path
    ) -> None:
        with OdfPackage(odf_manifest_missing_part) as package:
            errors = package.validate_structure()

        assert any(
            "Manifest entry 'content.xml' was not found in package" in e.description
            for e in errors
        )

    def test_duplicate_manifest_entries_are_reported(
        self,
        odf_duplicate_manifest_entry: Path,
    ) -> None:
        with OdfPackage(odf_duplicate_manifest_entry) as package:
            errors = package.validate_structure()

        assert any(
            "Duplicate manifest file-entry path 'content.xml'" in e.description
            for e in errors
        )

    def test_missing_root_manifest_entry_is_reported(self, odf_missing_root_entry: Path) -> None:
        with OdfPackage(odf_missing_root_entry) as package:
            errors = package.validate_structure()

        assert any("missing required root file-entry '/'" in e.description for e in errors)

    def test_root_media_type_mismatch_is_reported(
        self,
        odf_root_mimetype_mismatch: Path,
    ) -> None:
        with OdfPackage(odf_root_mimetype_mismatch) as package:
            errors = package.validate_structure()

        assert any(
            "Root manifest media-type does not match mimetype" in e.description
            for e in errors
        )

    def test_unlisted_xml_part_is_error_in_strict_mode(self, odf_unlisted_xml_part: Path) -> None:
        with OdfPackage(odf_unlisted_xml_part) as package:
            errors = package.validate_structure(strict=True)

        assert any(
            "is not declared in manifest.xml" in e.description
            and e.severity == ValidationSeverity.ERROR
            for e in errors
        )

    def test_unlisted_xml_part_is_warning_in_permissive_mode(
        self,
        odf_unlisted_xml_part: Path,
    ) -> None:
        with OdfPackage(odf_unlisted_xml_part) as package:
            errors = package.validate_structure(strict=False)

        assert any(
            "is not declared in manifest.xml" in e.description
            and e.severity == ValidationSeverity.WARNING
            for e in errors
        )
