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

    def test_manifest_version_is_exposed(self, minimal_odt: Path) -> None:
        with OdfPackage(minimal_odt) as package:
            assert package.manifest_version == "1.3"

    def test_manifest_encryption_metadata_is_exposed(
        self,
        minimal_odt_encrypted_structural: Path,
    ) -> None:
        with OdfPackage(minimal_odt_encrypted_structural) as package:
            content_entries = [
                entry
                for entry in package.manifest
                if entry.full_path.strip().lstrip("/") == "content.xml"
            ]
            assert content_entries
            content = content_entries[0]
            assert content.has_encryption_data
            assert content.encryption_algorithm_name
            assert content.encryption_key_derivation_name

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

    def test_missing_content_xml_is_reported(self, odf_missing_content_xml: Path) -> None:
        with OdfPackage(odf_missing_content_xml) as package:
            errors = package.validate_structure(strict=True)

        assert any("missing required member 'content.xml'" in e.description for e in errors)
        assert any("missing required manifest entry 'content.xml'" in e.description for e in errors)

    def test_manifest_declared_auxiliary_missing_part_is_reported(
        self,
        odf_aux_declared_missing_styles: Path,
    ) -> None:
        with OdfPackage(odf_aux_declared_missing_styles) as package:
            errors = package.validate_structure(strict=True)

        assert any(
            "Manifest entry 'styles.xml' was not found in package" in e.description
            for e in errors
        )

    def test_signature_manifest_missing_part_is_reported(
        self,
        odf_signature_manifest_missing_xml: Path,
    ) -> None:
        with OdfPackage(odf_signature_manifest_missing_xml) as package:
            errors = package.validate_structure(strict=True)

        assert any(
            "Manifest entry 'META-INF/documentsignatures.xml' was not found in package"
            in e.description
            for e in errors
        )
