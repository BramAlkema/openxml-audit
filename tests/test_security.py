"""Tests for OOXML security validation rules."""

from __future__ import annotations

import zipfile
from pathlib import Path

from openxml_audit import OpenXmlValidator, ValidationSeverity


def _read_zip_text(path: Path, member: str) -> str:
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read(member).decode("utf-8")


def _clone_zip(
    source: Path,
    target: Path,
    updates: dict[str, str | bytes],
) -> Path:
    with (
        zipfile.ZipFile(source, "r") as src,
        zipfile.ZipFile(
            target,
            "w",
            zipfile.ZIP_DEFLATED,
        ) as dst,
    ):
        written: set[str] = set()
        for info in src.infolist():
            payload = updates.get(info.filename)
            if payload is None:
                dst.writestr(info, src.read(info.filename))
                continue
            data = payload.encode("utf-8") if isinstance(payload, str) else payload
            dst.writestr(info.filename, data)
            written.add(info.filename)

        for name, payload in updates.items():
            if name in written:
                continue
            data = payload.encode("utf-8") if isinstance(payload, str) else payload
            dst.writestr(name, data)

    return target


def _append_relationship(xml: str, relationship_xml: str) -> str:
    return xml.replace("</Relationships>", f"{relationship_xml}\n</Relationships>")


def test_dangerous_uri_scheme_is_reported_as_error(
    minimal_pptx: Path,
    tmp_path: Path,
) -> None:
    rels = _read_zip_text(minimal_pptx, "ppt/slides/_rels/slide1.xml.rels")
    updated_rels = _append_relationship(
        rels,
        (
            '    <Relationship Id="rId9" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            'Target="javascript:alert(1)" TargetMode="External"/>'
        ),
    )
    package_path = _clone_zip(
        minimal_pptx,
        tmp_path / "dangerous-uri.pptx",
        {"ppt/slides/_rels/slide1.xml.rels": updated_rels},
    )

    result = OpenXmlValidator(schema_validation=False).validate(package_path)

    assert not result.is_valid
    assert any(
        error.id == "Sec_DangerousUri"
        and error.severity == ValidationSeverity.ERROR
        and "dangerous URI scheme" in error.description
        and error.part_uri == "/ppt/slides/_rels/slide1.xml.rels"
        for error in result.errors
    )


def test_external_https_relationship_warns_but_stays_valid(
    minimal_pptx: Path,
    tmp_path: Path,
) -> None:
    rels = _read_zip_text(minimal_pptx, "ppt/slides/_rels/slide1.xml.rels")
    updated_rels = _append_relationship(
        rels,
        (
            '    <Relationship Id="rId9" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            'Target="https://example.com" TargetMode="External"/>'
        ),
    )
    package_path = _clone_zip(
        minimal_pptx,
        tmp_path / "external-link.pptx",
        {"ppt/slides/_rels/slide1.xml.rels": updated_rels},
    )

    result = OpenXmlValidator(schema_validation=False).validate(package_path)

    assert result.is_valid
    assert any(
        error.id == "Sec_ExternalRelationship"
        and error.severity == ValidationSeverity.WARNING
        and "https://example.com" in error.description
        for error in result.errors
    )


def test_ssrf_like_relationship_target_is_reported_as_warning(
    minimal_pptx: Path,
    tmp_path: Path,
) -> None:
    rels = _read_zip_text(minimal_pptx, "ppt/slides/_rels/slide1.xml.rels")
    updated_rels = _append_relationship(
        rels,
        (
            '    <Relationship Id="rId9" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            'Target="http://169.254.169.254/latest/meta-data/" TargetMode="External"/>'
        ),
    )
    package_path = _clone_zip(
        minimal_pptx,
        tmp_path / "ssrf-target.pptx",
        {"ppt/slides/_rels/slide1.xml.rels": updated_rels},
    )

    result = OpenXmlValidator(schema_validation=False).validate(package_path)

    assert result.is_valid
    assert any(
        error.id == "Sec_SsrfUri"
        and error.severity == ValidationSeverity.WARNING
        and "metadata endpoint" in error.description
        for error in result.errors
    )


def test_active_content_element_is_reported_as_error(
    minimal_pptx: Path,
    tmp_path: Path,
) -> None:
    slide_xml = _read_zip_text(minimal_pptx, "ppt/slides/slide1.xml")
    updated_slide = slide_xml.replace(
        "        <p:spTree>\n",
        "        <p:spTree>\n            <p:oleObj/>\n",
    )
    package_path = _clone_zip(
        minimal_pptx,
        tmp_path / "active-element.pptx",
        {"ppt/slides/slide1.xml": updated_slide},
    )

    result = OpenXmlValidator(schema_validation=False).validate(package_path)

    assert not result.is_valid
    assert any(
        error.id == "Sec_ActiveContentElement"
        and error.severity == ValidationSeverity.ERROR
        and "oleObj" in error.description
        and error.part_uri == "/ppt/slides/slide1.xml"
        for error in result.errors
    )


def test_macro_enabled_main_content_type_is_reported_as_error(
    minimal_xlsx: Path,
    tmp_path: Path,
) -> None:
    content_types = _read_zip_text(minimal_xlsx, "[Content_Types].xml")
    updated_content_types = content_types.replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    )
    package_path = _clone_zip(
        minimal_xlsx,
        tmp_path / "macro-enabled.xlsm",
        {"[Content_Types].xml": updated_content_types},
    )

    result = OpenXmlValidator(schema_validation=False).validate(package_path)

    assert not result.is_valid
    assert any(
        error.id == "Sec_ActiveContentType"
        and error.severity == ValidationSeverity.ERROR
        and "macroEnabled" in error.description
        and error.part_uri == "/[Content_Types].xml"
        for error in result.errors
    )
