"""pytest configuration and fixtures for openxml_audit tests."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from openxml_audit import OpenXmlValidator
from tests.fixture_loader import FIXTURES_DIR


@pytest.fixture
def openxml_audit() -> OpenXmlValidator:
    """Provide an OpenXmlValidator instance."""
    return OpenXmlValidator()


@pytest.fixture
def tmp_pptx_path(tmp_path: Path) -> Path:
    """Provide a temporary path for PPTX files."""
    return tmp_path / "test.pptx"


def _is_xml_file(path: Path) -> bool:
    if path.name == "[Content_Types].xml":
        return True
    return path.suffix in {".xml", ".rels"}


def _build_package_from_dir(
    source_dir: Path,
    output_path: Path,
    *,
    validate_xml: bool = True,
) -> Path:
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_dir():
                continue
            rel_path = file_path.relative_to(source_dir).as_posix()
            data = file_path.read_bytes()
            if validate_xml and _is_xml_file(file_path):
                # Validate XML fixtures up front.
                etree.fromstring(data)
            zf.writestr(rel_path, data)

    output_path.write_bytes(buffer.getvalue())
    return output_path


@pytest.fixture
def minimal_pptx(tmp_path: Path) -> Path:
    """Create a minimal valid PPTX file."""
    pptx_path = tmp_path / "minimal.pptx"
    return _build_package_from_dir(FIXTURES_DIR / "pptx" / "minimal", pptx_path)


@pytest.fixture
def minimal_docx(tmp_path: Path) -> Path:
    """Create a minimal valid DOCX file."""
    docx_path = tmp_path / "minimal.docx"
    return _build_package_from_dir(FIXTURES_DIR / "docx" / "minimal", docx_path)


@pytest.fixture
def minimal_xlsx(tmp_path: Path) -> Path:
    """Create a minimal valid XLSX file."""
    xlsx_path = tmp_path / "minimal.xlsx"
    return _build_package_from_dir(FIXTURES_DIR / "xlsx" / "minimal", xlsx_path)


@pytest.fixture
def docx_missing_body(tmp_path: Path) -> Path:
    """Create a DOCX missing the body element."""
    docx_path = tmp_path / "missing_body.docx"
    return _build_package_from_dir(FIXTURES_DIR / "docx" / "missing_body", docx_path)


@pytest.fixture
def docx_missing_styles_with_effects(tmp_path: Path) -> Path:
    """Create a DOCX missing stylesWithEffects relationship."""
    docx_path = tmp_path / "missing_styles_with_effects.docx"
    return _build_package_from_dir(
        FIXTURES_DIR / "docx" / "missing_styles_with_effects",
        docx_path,
    )


@pytest.fixture
def xlsx_missing_sheet_rel(tmp_path: Path) -> Path:
    """Create an XLSX with a missing sheet relationship."""
    xlsx_path = tmp_path / "missing_sheet_rel.xlsx"
    return _build_package_from_dir(
        FIXTURES_DIR / "xlsx" / "missing_sheet_rel",
        xlsx_path,
    )


@pytest.fixture
def xlsx_missing_shared_strings(tmp_path: Path) -> Path:
    """Create an XLSX missing sharedStrings.xml but referencing shared strings."""
    xlsx_path = tmp_path / "missing_shared_strings.xlsx"
    return _build_package_from_dir(
        FIXTURES_DIR / "xlsx" / "missing_shared_strings",
        xlsx_path,
    )


@pytest.fixture
def invalid_pptx_missing_presentation(tmp_path: Path) -> Path:
    """Create invalid PPTX missing presentation.xml."""
    pptx_path = tmp_path / "missing_presentation.pptx"
    return _build_package_from_dir(
        FIXTURES_DIR / "pptx" / "missing_presentation",
        pptx_path,
    )


@pytest.fixture
def not_a_zip(tmp_path: Path) -> Path:
    """Create a file that is not a valid ZIP."""
    path = tmp_path / "not_a_zip.pptx"
    path.write_text("This is not a ZIP file")
    return path


def _build_odf_fixture(
    tmp_path: Path,
    fixture_relpath: str,
    *,
    filename: str,
    validate_xml: bool = True,
) -> Path:
    source_dir = FIXTURES_DIR / "odf" / fixture_relpath
    return _build_package_from_dir(
        source_dir,
        tmp_path / filename,
        validate_xml=validate_xml,
    )


@pytest.fixture
def minimal_odt(tmp_path: Path) -> Path:
    """Create a minimal valid ODT file."""
    return _build_odf_fixture(
        tmp_path,
        "valid/minimal_odt",
        filename="minimal.odt",
    )


@pytest.fixture
def minimal_ods(tmp_path: Path) -> Path:
    """Create a minimal valid ODS file."""
    return _build_odf_fixture(
        tmp_path,
        "valid/minimal_ods",
        filename="minimal.ods",
    )


@pytest.fixture
def minimal_odp(tmp_path: Path) -> Path:
    """Create a minimal valid ODP file."""
    return _build_odf_fixture(
        tmp_path,
        "valid/minimal_odp",
        filename="minimal.odp",
    )


@pytest.fixture
def odf_missing_mimetype(tmp_path: Path) -> Path:
    """Create an ODF package missing the mimetype entry."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/missing_mimetype",
        filename="missing_mimetype.odt",
    )


@pytest.fixture
def odf_invalid_mimetype(tmp_path: Path) -> Path:
    """Create an ODF package with an invalid mimetype value."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/invalid_mimetype",
        filename="invalid_mimetype.odt",
    )


@pytest.fixture
def odf_missing_manifest(tmp_path: Path) -> Path:
    """Create an ODF package missing META-INF/manifest.xml."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/missing_manifest",
        filename="missing_manifest.odt",
    )


@pytest.fixture
def odf_malformed_manifest(tmp_path: Path) -> Path:
    """Create an ODF package with malformed manifest XML."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/malformed_manifest",
        filename="malformed_manifest.odt",
        validate_xml=False,
    )


@pytest.fixture
def odf_manifest_missing_part(tmp_path: Path) -> Path:
    """Create an ODF package where manifest references a missing part."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/manifest_missing_part",
        filename="manifest_missing_part.odt",
    )


@pytest.fixture
def odf_unlisted_xml_part(tmp_path: Path) -> Path:
    """Create an ODF package with XML content not declared in manifest."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/unlisted_xml_part",
        filename="unlisted_xml_part.odt",
    )


@pytest.fixture
def odf_duplicate_manifest_entry(tmp_path: Path) -> Path:
    """Create an ODF package with duplicate manifest file-entry paths."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/duplicate_manifest_entry",
        filename="duplicate_manifest_entry.odt",
    )


@pytest.fixture
def odf_missing_root_entry(tmp_path: Path) -> Path:
    """Create an ODF package with no root ('/') manifest file-entry."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/missing_root_entry",
        filename="missing_root_entry.odt",
    )


@pytest.fixture
def odf_root_mimetype_mismatch(tmp_path: Path) -> Path:
    """Create an ODF package where root manifest media-type mismatches mimetype."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/root_mimetype_mismatch",
        filename="root_mimetype_mismatch.odt",
    )


@pytest.fixture
def odf_broken_content_xml(tmp_path: Path) -> Path:
    """Create an ODF package with malformed content.xml."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/broken_content_xml",
        filename="broken_content_xml.odt",
        validate_xml=False,
    )


@pytest.fixture
def odf_missing_content_xml(tmp_path: Path) -> Path:
    """Create an ODF package missing required content.xml."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/missing_content_xml",
        filename="missing_content_xml.odt",
    )


@pytest.fixture
def odf_content_body_mismatch(tmp_path: Path) -> Path:
    """Create an ODF package where content.xml body type mismatches mimetype."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/content_body_mismatch",
        filename="content_body_mismatch.odt",
    )


@pytest.fixture
def odf_content_root_mismatch(tmp_path: Path) -> Path:
    """Create an ODF package where content.xml root element is invalid for content part."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/content_root_mismatch",
        filename="content_root_mismatch.odt",
    )
