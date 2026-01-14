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


def _build_package_from_dir(source_dir: Path, output_path: Path) -> Path:
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_dir():
                continue
            rel_path = file_path.relative_to(source_dir).as_posix()
            data = file_path.read_bytes()
            if _is_xml_file(file_path):
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
