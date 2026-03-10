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
def minimal_odt_v12(tmp_path: Path) -> Path:
    """Create a minimal valid ODT file with ODF 1.2 version markers."""
    return _build_odf_fixture(
        tmp_path,
        "valid/minimal_odt_v12",
        filename="minimal_v12.odt",
    )


@pytest.fixture
def minimal_odt_v14(tmp_path: Path) -> Path:
    """Create a minimal valid ODT file with ODF 1.4 version markers."""
    return _build_odf_fixture(
        tmp_path,
        "valid/minimal_odt_v14",
        filename="minimal_v14.odt",
    )


@pytest.fixture
def minimal_odt_with_styles(tmp_path: Path) -> Path:
    """Create a minimal valid ODT file with content.xml and styles.xml."""
    return _build_odf_fixture(
        tmp_path,
        "valid/minimal_odt_with_styles",
        filename="minimal_with_styles.odt",
    )


@pytest.fixture
def minimal_odt_text_style_reference_with_styles(tmp_path: Path) -> Path:
    """Create a minimal valid ODT with text:style-name and styles.xml companion."""
    return _build_odf_fixture(
        tmp_path,
        "valid/text_style_reference_with_styles",
        filename="text_style_reference_with_styles.odt",
    )


@pytest.fixture
def minimal_odp_master_page_resolved(tmp_path: Path) -> Path:
    """Create a minimal valid ODP with draw:master-page-name resolved by styles.xml."""
    return _build_odf_fixture(
        tmp_path,
        "valid/presentation_master_page_resolved",
        filename="presentation_master_page_resolved.odp",
    )


@pytest.fixture
def minimal_odt_signed_stub(tmp_path: Path) -> Path:
    """Create a minimal valid ODT-like package with signature metadata stub."""
    return _build_odf_fixture(
        tmp_path,
        "valid/signed_stub_odt",
        filename="signed_stub.odt",
    )


@pytest.fixture
def minimal_odt_encrypted_stub(tmp_path: Path) -> Path:
    """Create a minimal valid ODT-like package with encryption metadata stub."""
    return _build_odf_fixture(
        tmp_path,
        "valid/encrypted_stub_odt",
        filename="encrypted_stub.odt",
    )


@pytest.fixture
def minimal_odt_signed_structural(tmp_path: Path) -> Path:
    """Create a minimal ODT with structural signature metadata."""
    return _build_odf_fixture(
        tmp_path,
        "valid/signed_structural_odt",
        filename="signed_structural.odt",
    )


@pytest.fixture
def minimal_odt_encrypted_structural(tmp_path: Path) -> Path:
    """Create a minimal ODT with structural encryption metadata."""
    return _build_odf_fixture(
        tmp_path,
        "valid/encrypted_structural_odt",
        filename="encrypted_structural.odt",
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


@pytest.fixture
def odf_aux_declared_missing_styles(tmp_path: Path) -> Path:
    """Create an ODF package with manifest-declared styles.xml missing from ZIP."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/aux_declared_missing_styles",
        filename="aux_declared_missing_styles.odt",
    )


@pytest.fixture
def odf_aux_invalid_styles_xml(tmp_path: Path) -> Path:
    """Create an ODF package with malformed styles.xml declared by manifest."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/aux_invalid_styles_xml",
        filename="aux_invalid_styles_xml.odt",
        validate_xml=False,
    )


@pytest.fixture
def odf_manifest_content_bad_media_type(tmp_path: Path) -> Path:
    """Create an ODF package with content.xml manifest media-type mismatch."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/manifest_content_bad_media_type",
        filename="manifest_content_bad_media_type.odt",
    )


@pytest.fixture
def odf_text_style_reference_missing_styles(tmp_path: Path) -> Path:
    """Create a text ODF package with style references and no styles.xml companion."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/text_style_reference_missing_styles",
        filename="text_style_reference_missing_styles.odt",
    )


@pytest.fixture
def odf_spreadsheet_duplicate_table_names(tmp_path: Path) -> Path:
    """Create a spreadsheet ODF package with duplicate table:name values."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/spreadsheet_duplicate_table_names",
        filename="spreadsheet_duplicate_table_names.ods",
    )


@pytest.fixture
def odf_presentation_duplicate_page_names(tmp_path: Path) -> Path:
    """Create a presentation ODF package with duplicate draw:name values."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/presentation_duplicate_page_names",
        filename="presentation_duplicate_page_names.odp",
    )


@pytest.fixture
def odf_presentation_master_page_missing_styles(tmp_path: Path) -> Path:
    """Create a presentation ODF package that references master pages without styles.xml."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/presentation_master_page_missing_styles",
        filename="presentation_master_page_missing_styles.odp",
    )


@pytest.fixture
def odf_presentation_master_page_unresolved(tmp_path: Path) -> Path:
    """Create a presentation ODF package with unresolved draw:master-page-name."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/presentation_master_page_unresolved",
        filename="presentation_master_page_unresolved.odp",
    )


@pytest.fixture
def odf_signature_manifest_missing_xml(tmp_path: Path) -> Path:
    """Create an ODF package with missing META-INF/documentsignatures.xml part."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/signature_manifest_missing_xml",
        filename="signature_manifest_missing_xml.odt",
    )


@pytest.fixture
def odf_signature_bad_root(tmp_path: Path) -> Path:
    """Create an ODF package with invalid documentsignatures.xml root."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/signature_bad_root",
        filename="signature_bad_root.odt",
    )


@pytest.fixture
def odf_signature_bad_media_type(tmp_path: Path) -> Path:
    """Create an ODF package with wrong signature entry media-type."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/signature_bad_media_type",
        filename="signature_bad_media_type.odt",
    )


@pytest.fixture
def odf_signature_missing_signedinfo(tmp_path: Path) -> Path:
    """Create an ODF package with ds:Signature missing SignedInfo/Reference."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/signature_missing_signedinfo",
        filename="signature_missing_signedinfo.odt",
    )


@pytest.fixture
def odf_encrypted_missing_key_derivation(tmp_path: Path) -> Path:
    """Create an ODF package with encryption-data missing key-derivation."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/encrypted_missing_key_derivation",
        filename="encrypted_missing_key_derivation.odt",
    )


@pytest.fixture
def odf_encrypted_root_entry_encrypted(tmp_path: Path) -> Path:
    """Create an ODF package with invalid encryption-data on root entry '/'."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/encrypted_root_entry_encrypted",
        filename="encrypted_root_entry_encrypted.odt",
    )


@pytest.fixture
def odf_encrypted_checksum_partial(tmp_path: Path) -> Path:
    """Create an ODF package with incomplete encryption checksum metadata."""
    return _build_odf_fixture(
        tmp_path,
        "invalid/encrypted_checksum_partial",
        filename="encrypted_checksum_partial.odt",
    )


# ── new semantic rule fixtures ───────────────────────────────────────


@pytest.fixture
def odf_meta_missing_office_meta(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/meta_missing_office_meta", filename="meta_missing.odt")


@pytest.fixture
def odf_settings_missing_office_settings(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/settings_missing_office_settings", filename="settings_missing.odt")


@pytest.fixture
def odf_font_face_missing_svg_family(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/font_face_missing_svg_family", filename="font_face_bad.odt")


@pytest.fixture
def odf_style_parent_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/style_parent_unresolved", filename="parent_unresolved.odt")


@pytest.fixture
def odf_style_data_style_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/style_data_style_unresolved", filename="data_style_unresolved.ods")


@pytest.fixture
def odf_style_list_style_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/style_list_style_unresolved", filename="list_style_unresolved.odt")


@pytest.fixture
def odf_master_page_layout_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/master_page_layout_unresolved", filename="master_layout_unresolved.odt")


@pytest.fixture
def odf_text_list_style_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/text_list_style_unresolved", filename="text_list_unresolved.odt")


@pytest.fixture
def odf_text_bookmark_ref_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/text_bookmark_ref_unresolved", filename="bookmark_unresolved.odt")


@pytest.fixture
def odf_spreadsheet_named_range_bad_table(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/spreadsheet_named_range_bad_table", filename="named_range_bad.ods")


@pytest.fixture
def odf_spreadsheet_column_overflow(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/spreadsheet_column_overflow", filename="column_overflow.ods")


@pytest.fixture
def odf_presentation_no_pages(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/presentation_no_pages", filename="no_pages.odp")


@pytest.fixture
def odf_presentation_page_layout_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/presentation_page_layout_unresolved", filename="page_layout_unresolved.odp")


@pytest.fixture
def odf_embedded_object_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/embedded_object_unresolved", filename="object_unresolved.odt")


@pytest.fixture
def odf_image_href_unresolved(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/image_href_unresolved", filename="image_unresolved.odt")


@pytest.fixture
def odf_meta_bad_statistics(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "invalid/meta_bad_statistics", filename="meta_bad_stats.odt")


@pytest.fixture
def odf_valid_meta(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "valid/minimal_odt_with_meta", filename="valid_meta.odt")


@pytest.fixture
def odf_valid_font_face(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "valid/font_face_with_svg_family", filename="valid_font.odt")


@pytest.fixture
def odf_valid_style_parent(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "valid/style_parent_resolved", filename="valid_parent.odt")


@pytest.fixture
def odf_valid_named_range(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "valid/spreadsheet_named_range_valid", filename="valid_named_range.ods")


@pytest.fixture
def odf_valid_bookmark(tmp_path: Path) -> Path:
    return _build_odf_fixture(tmp_path, "valid/text_bookmark_resolved", filename="valid_bookmark.odt")
