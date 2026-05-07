"""Integration tests for Word and Excel validators."""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from openxml_audit import OpenXmlValidator, ValidationErrorType
from openxml_audit.namespaces import CONTENT_TYPES, REL_WEB_SETTINGS, RELATIONSHIPS


def _remove_docx_web_settings(source: Path, target: Path) -> Path:
    with (
        zipfile.ZipFile(source) as src,
        zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst,
    ):
        for name in src.namelist():
            if name == "word/webSettings.xml":
                continue

            content = src.read(name)
            if name == "word/_rels/document.xml.rels":
                root = etree.fromstring(content)
                for rel in root.findall(f"{{{RELATIONSHIPS}}}Relationship"):
                    if rel.get("Type") == REL_WEB_SETTINGS:
                        root.remove(rel)
                content = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
            elif name == "[Content_Types].xml":
                root = etree.fromstring(content)
                for override in root.findall(f"{{{CONTENT_TYPES}}}Override"):
                    if override.get("PartName") == "/word/webSettings.xml":
                        root.remove(override)
                content = etree.tostring(root, xml_declaration=True, encoding="UTF-8")

            dst.writestr(name, content)

    return target


def test_minimal_docx_valid(minimal_docx: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(minimal_docx)

    assert result.is_valid


def test_docx_missing_body_reports_error(docx_missing_body: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(docx_missing_body)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SCHEMA and "body" in e.description
        for e in result.errors
    )


def test_docx_missing_styles_with_effects_is_valid(
    docx_missing_styles_with_effects: Path,
) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=True)
    result = validator.validate(docx_missing_styles_with_effects)

    assert result.is_valid


def test_docx_without_web_settings_is_valid(
    tmp_path: Path,
    docx_missing_styles_with_effects: Path,
) -> None:
    docx = _remove_docx_web_settings(
        docx_missing_styles_with_effects,
        tmp_path / "without_web_settings.docx",
    )
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=True)
    result = validator.validate(docx)

    assert result.is_valid, [e.description for e in result.errors]


def test_minimal_xlsx_valid(minimal_xlsx: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(minimal_xlsx)

    assert result.is_valid


def test_xlsx_missing_sheet_relationship_reports_error(xlsx_missing_sheet_rel: Path) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(xlsx_missing_sheet_rel)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SEMANTIC and "relationship" in e.description
        for e in result.errors
    )


def test_xlsx_missing_shared_strings_reports_error(
    xlsx_missing_shared_strings: Path,
) -> None:
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=True)
    result = validator.validate(xlsx_missing_shared_strings)

    assert not result.is_valid
    assert any(
        e.error_type == ValidationErrorType.SEMANTIC
        and "sharedStrings" in e.description
        for e in result.errors
    )
