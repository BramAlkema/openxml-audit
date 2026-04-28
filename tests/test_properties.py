"""Tests for extended, core, and custom properties validation."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from openxml_audit import OpenXmlValidator


def _build_pkg(tmp_path: Path, name: str, files: dict[str, str]) -> Path:
    """Build a .docx/.pptx from a dict of path->content."""
    path = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p, content in files.items():
            zf.writestr(p, content)
    path.write_bytes(buf.getvalue())
    return path


_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  {extra}
</Relationships>"""

_CT = """\
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  {extra}
</Types>"""

_DOC = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body/>
</w:document>"""

_DOC_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""


def _base_files() -> dict[str, str]:
    return {
        "_rels/.rels": _RELS.format(extra=""),
        "[Content_Types].xml": _CT.format(extra=""),
        "word/document.xml": _DOC,
        "word/_rels/document.xml.rels": _DOC_RELS,
    }


class TestCoreProperties:

    def test_valid_core_properties(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        )
        files["[Content_Types].xml"] = _CT.format(
            extra='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        )
        files["docProps/core.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test</dc:title>
  <dc:creator>Author</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
  <cp:revision>1</cp:revision>
</cp:coreProperties>"""
        pkg = _build_pkg(tmp_path, "valid_core.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        core_errors = [e for e in result.errors if "core" in e.description.lower() or "coreProperties" in e.description]
        assert core_errors == [], [e.description for e in core_errors]

    def test_core_bad_root(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        )
        files["[Content_Types].xml"] = _CT.format(
            extra='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        )
        files["docProps/core.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<wrong xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>"""
        pkg = _build_pkg(tmp_path, "bad_core.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("coreProperties" in d for d in descs), descs

    def test_core_unexpected_element(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        )
        files["[Content_Types].xml"] = _CT.format(
            extra='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        )
        files["docProps/core.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
  <cp:bogus>test</cp:bogus>
</cp:coreProperties>"""
        pkg = _build_pkg(tmp_path, "unexpected_core.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("Unexpected" in d for d in descs), descs


class TestExtendedProperties:

    def test_valid_extended_properties(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        )
        files["[Content_Types].xml"] = _CT.format(
            extra='<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        )
        files["docProps/app.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <Pages>1</Pages>
  <Words>100</Words>
  <ScaleCrop>false</ScaleCrop>
</Properties>"""
        pkg = _build_pkg(tmp_path, "valid_app.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        app_errors = [e for e in result.errors if "extended" in e.description.lower() or "Extended" in e.description]
        assert app_errors == [], [e.description for e in app_errors]

    def test_extended_bad_int(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        )
        files["docProps/app.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Pages>not_a_number</Pages>
</Properties>"""
        pkg = _build_pkg(tmp_path, "bad_int.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("integer" in d for d in descs), descs

    def test_extended_bad_bool(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        )
        files["docProps/app.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <ScaleCrop>maybe</ScaleCrop>
</Properties>"""
        pkg = _build_pkg(tmp_path, "bad_bool.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("boolean" in d for d in descs), descs

    def test_extended_vector_size_mismatch(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        )
        files["docProps/app.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TitlesOfParts>
    <vt:vector size="3" baseType="lpstr">
      <vt:lpstr>Sheet1</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
</Properties>"""
        pkg = _build_pkg(tmp_path, "bad_vector.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("size=3" in d and "1 children" in d for d in descs), descs


class TestCustomProperties:

    def test_valid_custom_properties(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        )
        files["docProps/custom.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Project">
    <vt:lpwstr>Test</vt:lpwstr>
  </property>
</Properties>"""
        pkg = _build_pkg(tmp_path, "valid_custom.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        custom_errors = [e for e in result.errors if "custom" in e.description.lower() or "Custom" in e.description]
        assert custom_errors == [], [e.description for e in custom_errors]

    def test_custom_duplicate_name(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        )
        files["docProps/custom.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Dup">
    <vt:lpwstr>A</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Dup">
    <vt:lpwstr>B</vt:lpwstr>
  </property>
</Properties>"""
        pkg = _build_pkg(tmp_path, "dup_name.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("Duplicate" in d and "Dup" in d for d in descs), descs

    def test_custom_pid_below_2(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        )
        files["docProps/custom.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="1" name="Bad">
    <vt:lpwstr>X</vt:lpwstr>
  </property>
</Properties>"""
        pkg = _build_pkg(tmp_path, "bad_pid.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any(">= 2" in d for d in descs), descs

    def test_custom_no_value(self, tmp_path: Path) -> None:
        files = _base_files()
        files["_rels/.rels"] = _RELS.format(
            extra='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        )
        files["docProps/custom.xml"] = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Empty"/>
</Properties>"""
        pkg = _build_pkg(tmp_path, "no_value.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("no value" in d for d in descs), descs


class TestStylesWithEffects:

    def _word_files_with_swe(self, swe_xml: str, styles_xml: str | None = None) -> dict[str, str]:
        doc_rels = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/>
  {extra}
</Relationships>"""
        extra_rels = ""
        if styles_xml:
            extra_rels = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'

        files = _base_files()
        files["word/_rels/document.xml.rels"] = doc_rels.format(extra=extra_rels)
        files["[Content_Types].xml"] = _CT.format(
            extra='<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/>'
        )
        files["word/stylesWithEffects.xml"] = swe_xml
        if styles_xml:
            files["word/styles.xml"] = styles_xml
        return files

    def test_valid_styles_with_effects(self, tmp_path: Path) -> None:
        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"""
        files = self._word_files_with_swe(swe)
        pkg = _build_pkg(tmp_path, "valid_swe.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        swe_errors = [e for e in result.errors if "stylesWithEffects" in e.description]
        assert swe_errors == [], [e.description for e in swe_errors]

    def test_swe_bad_root(self, tmp_path: Path) -> None:
        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:wrong xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"""
        files = self._word_files_with_swe(swe)
        pkg = _build_pkg(tmp_path, "bad_root_swe.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("root" in d.lower() and "styles" in d.lower() for d in descs), descs

    def test_swe_duplicate_style_id(self, tmp_path: Path) -> None:
        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Dup"/>
  <w:style w:type="paragraph" w:styleId="Dup"/>
</w:styles>"""
        files = self._word_files_with_swe(swe)
        pkg = _build_pkg(tmp_path, "dup_swe.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("duplicate" in d.lower() and "Dup" in d for d in descs), descs

    def test_swe_invalid_type(self, tmp_path: Path) -> None:
        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="bogus" w:styleId="S1"/>
</w:styles>"""
        files = self._word_files_with_swe(swe)
        pkg = _build_pkg(tmp_path, "bad_type_swe.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        descs = [e.description for e in result.errors]
        assert any("invalid" in d.lower() and "bogus" in d for d in descs), descs

    def test_swe_orphaned_styles(self, tmp_path: Path) -> None:
        """Styles in stylesWithEffects but not in styles.xml are an ERROR."""
        from openxml_audit.errors import ValidationSeverity

        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"/>
  <w:style w:type="paragraph" w:styleId="OrphanStyle"/>
</w:styles>"""
        styles = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>"""
        files = self._word_files_with_swe(swe, styles_xml=styles)
        pkg = _build_pkg(tmp_path, "orphan_swe.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        matching = [e for e in result.errors if "OrphanStyle" in e.description]
        assert matching, [e.description for e in result.errors]
        assert all(
            e.severity == ValidationSeverity.ERROR for e in matching
        ), [(e.severity, e.description) for e in matching]

    def test_swe_styles_missing_from_effects(self, tmp_path: Path) -> None:
        """Styles in styles.xml but not in stylesWithEffects are an ERROR.

        This is the python-docx failure mode: tools that modify styles.xml
        without updating stylesWithEffects produce documents that Word
        flags as unreadable content.
        """
        from openxml_audit.errors import ValidationSeverity

        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>"""
        styles = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"/>
  <w:style w:type="character" w:styleId="Hyperlink"/>
  <w:style w:type="paragraph" w:styleId="Header"/>
</w:styles>"""
        files = self._word_files_with_swe(swe, styles_xml=styles)
        pkg = _build_pkg(tmp_path, "missing_in_swe.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        matching = [
            e for e in result.errors
            if "stylesWithEffects" in e.description and "Hyperlink" in e.description
        ]
        assert matching, [e.description for e in result.errors]
        assert all(
            e.severity == ValidationSeverity.ERROR for e in matching
        ), [(e.severity, e.description) for e in matching]
        assert any("Header" in e.description for e in matching)

    def test_swe_doc_defaults_differ(self, tmp_path: Path) -> None:
        """Differing docDefaults between styles.xml and stylesWithEffects → ERROR."""
        from openxml_audit.errors import ValidationSeverity

        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>"""
        styles = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>"""
        files = self._word_files_with_swe(swe, styles_xml=styles)
        pkg = _build_pkg(tmp_path, "doc_defaults_differ.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        matching = [e for e in result.errors if "docDefaults" in e.description]
        assert matching, [e.description for e in result.errors]
        assert all(e.severity == ValidationSeverity.ERROR for e in matching)

    def test_swe_doc_defaults_one_sided(self, tmp_path: Path) -> None:
        """docDefaults present in only one of the two parts → ERROR."""
        from openxml_audit.errors import ValidationSeverity

        swe = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>"""
        styles = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"/>
</w:styles>"""
        files = self._word_files_with_swe(swe, styles_xml=styles)
        pkg = _build_pkg(tmp_path, "doc_defaults_one_sided.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        matching = [
            e for e in result.errors
            if "docDefaults" in e.description and "missing" in e.description
        ]
        assert matching, [e.description for e in result.errors]
        assert all(e.severity == ValidationSeverity.ERROR for e in matching)

    def test_swe_consistent_with_styles(self, tmp_path: Path) -> None:
        """Identical styles.xml and stylesWithEffects produce no consistency errors."""
        body = """\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr>
  </w:style>
</w:styles>"""
        files = self._word_files_with_swe(body, styles_xml=body)
        pkg = _build_pkg(tmp_path, "consistent.docx", files)
        result = OpenXmlValidator(schema_validation=False).validate(pkg)
        consistency_errors = [
            e for e in result.errors
            if any(
                marker in e.description
                for marker in ("docDefaults", "differ", "not in styles.xml", "not in stylesWithEffects")
            )
        ]
        assert consistency_errors == [], [e.description for e in consistency_errors]
