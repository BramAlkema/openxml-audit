"""Tests for OPC package handling."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from openxml_audit.errors import PackageValidationError
from openxml_audit.package import ContentTypes, OpenXmlPackage
from tests.fixture_loader import load_fixture_bytes


class TestContentTypes:
    """Tests for ContentTypes parsing."""

    def test_parse_default_content_types(self) -> None:
        """Test parsing default content types."""
        xml = load_fixture_bytes("content_types", "defaults.xml")

        ct = ContentTypes.from_xml(xml)

        assert ct.get_content_type("/test.rels") == "application/vnd.openxmlformats-package.relationships+xml"
        assert ct.get_content_type("/test.xml") == "application/xml"

    def test_parse_override_content_types(self) -> None:
        """Test parsing override content types."""
        xml = load_fixture_bytes("content_types", "override.xml")

        ct = ContentTypes.from_xml(xml)

        # Override takes precedence
        assert ct.get_content_type("/ppt/presentation.xml") == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        # Default still works for other xml files
        assert ct.get_content_type("/other.xml") == "application/xml"

    def test_unknown_extension(self) -> None:
        """Test getting content type for unknown extension."""
        xml = load_fixture_bytes("content_types", "empty.xml")

        ct = ContentTypes.from_xml(xml)
        assert ct.get_content_type("/test.unknown") is None


class TestOpenXmlPackage:
    """Tests for OpenXmlPackage."""

    def test_open_valid_package(self, minimal_pptx: Path) -> None:
        """Test opening a valid PPTX package."""
        with OpenXmlPackage(minimal_pptx) as package:
            parts = list(package.list_parts())
            assert "/ppt/presentation.xml" in parts

    def test_package_closes_on_exit(self, minimal_pptx: Path) -> None:
        """Test package closes when exiting context."""
        with OpenXmlPackage(minimal_pptx) as package:
            # Package is open inside context
            parts = list(package.list_parts())
            assert len(parts) > 0

        # After exiting, _zip should be None
        assert package._zip is None

    def test_get_main_document_uri(self, minimal_pptx: Path) -> None:
        """Test getting main document URI."""
        with OpenXmlPackage(minimal_pptx) as package:
            main_doc = package.get_main_document_uri()
            assert main_doc == "/ppt/presentation.xml"

    def test_has_part(self, minimal_pptx: Path) -> None:
        """Test checking for part existence."""
        with OpenXmlPackage(minimal_pptx) as package:
            assert package.has_part("/ppt/presentation.xml")
            assert package.has_part("ppt/presentation.xml")  # Without leading slash
            assert not package.has_part("/nonexistent.xml")

    def test_get_part_content(self, minimal_pptx: Path) -> None:
        """Test getting part content."""
        with OpenXmlPackage(minimal_pptx) as package:
            content = package.get_part_content("/ppt/presentation.xml")
            assert content is not None
            assert b"<p:presentation" in content

    def test_get_part_xml(self, minimal_pptx: Path) -> None:
        """Test getting part as parsed XML."""
        with OpenXmlPackage(minimal_pptx) as package:
            xml = package.get_part_xml("/ppt/presentation.xml")
            assert xml is not None
            assert "presentation" in xml.tag

    def test_package_relationships(self, minimal_pptx: Path) -> None:
        """Test accessing package relationships."""
        with OpenXmlPackage(minimal_pptx) as package:
            rels = list(package.relationships)
            assert len(rels) >= 1

            # Should have officeDocument relationship
            office_doc_rels = [r for r in rels if "officeDocument" in r.type]
            assert len(office_doc_rels) == 1

    def test_validate_structure(self, minimal_pptx: Path) -> None:
        """Test structure validation on valid package."""
        with OpenXmlPackage(minimal_pptx) as package:
            errors = package.validate_structure()
            # A minimal valid PPTX should have few or no errors
            error_msgs = [e.description for e in errors]
            assert "Cannot open file as ZIP" not in error_msgs
            assert "[Content_Types].xml not found" not in error_msgs

    def test_invalid_zip_raises_error(self, not_a_zip: Path) -> None:
        """Test that invalid ZIP raises PackageValidationError."""
        with pytest.raises(PackageValidationError) as exc_info:
            with OpenXmlPackage(not_a_zip):
                pass

        assert any("ZIP" in e.description for e in exc_info.value.errors)

    def test_missing_presentation_detected(
        self, invalid_pptx_missing_presentation: Path
    ) -> None:
        """Test that missing presentation.xml is detected."""
        # Invalid package doesn't necessarily raise on open - it collects errors
        with OpenXmlPackage(invalid_pptx_missing_presentation) as package:
            errors = package.validate_structure()

        # Should detect missing main document or missing officeDocument relationship
        assert len(errors) > 0
