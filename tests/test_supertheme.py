"""Tests for supertheme (themeVariantManager) validation."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from openxml_audit import OpenXmlValidator

# Shared XML fragments for building test .thmx packages

_CONTENT_TYPES_TEMPLATE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {extra_overrides}
</Types>"""

_PACKAGE_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  {extra_rels}
</Relationships>"""

_PRESENTATION_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

_PRESENTATION_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""

_SLIDE_MASTER_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:sldMaster>"""

_SLIDE_MASTER_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

_THEME_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
  <a:themeElements>
    <a:clrScheme name="Test Colors">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Test Fonts">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Test Format">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

_VARIANT_MANAGER_TEMPLATE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<t:themeVariantManager xmlns:t="http://schemas.microsoft.com/office/thememl/2012/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  {content}
</t:themeVariantManager>"""

_VARIANT_MANAGER_RELS_TEMPLATE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {rels}
</Relationships>"""

_VARIANT_PRESENTATION_RELS_TEMPLATE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""


def _build_thmx(
    tmp_path: Path,
    name: str,
    files: dict[str, str],
) -> Path:
    """Build a .thmx ZIP from a dict of path→content pairs."""
    thmx_path = tmp_path / name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, content in files.items():
            zf.writestr(path, content)
    thmx_path.write_bytes(buf.getvalue())
    return thmx_path


def _base_pptx_files() -> dict[str, str]:
    """Return the base PPTX files needed for a valid package."""
    return {
        "_rels/.rels": _PACKAGE_RELS.format(extra_rels=""),
        "[Content_Types].xml": _CONTENT_TYPES_TEMPLATE.format(extra_overrides=""),
        "ppt/presentation.xml": _PRESENTATION_XML,
        "ppt/_rels/presentation.xml.rels": _PRESENTATION_RELS,
        "ppt/slideMasters/slideMaster1.xml": _SLIDE_MASTER_XML,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": _SLIDE_MASTER_RELS,
        "ppt/theme/theme1.xml": _THEME_XML,
    }


def _supertheme_files(
    *,
    variant_count: int = 2,
    manager_content: str | None = None,
    manager_rels: str | None = None,
    include_variant_parts: bool = True,
    variant_theme_xml: str = _THEME_XML,
) -> dict[str, str]:
    """Build supertheme files to merge with base PPTX files."""
    files: dict[str, str] = {}

    # Package-level rel to themeVariantManager
    pkg_rels = _PACKAGE_RELS.format(
        extra_rels=(
            '<Relationship Id="rId2" '
            'Type="http://schemas.microsoft.com/office/2011/relationships/themeVariantManager" '
            'Target="themeVariants/themeVariantManager.xml"/>'
        )
    )
    files["_rels/.rels"] = pkg_rels

    # Content types with variant manager
    extra_ct = (
        '<Override PartName="/themeVariants/themeVariantManager.xml" '
        'ContentType="application/vnd.ms-powerpoint.themeVariantManager+xml"/>'
    )
    for i in range(1, variant_count + 1):
        extra_ct += (
            f'\n  <Override PartName="/themeVariants/variant{i}/theme/theme/theme1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        )
        extra_ct += (
            f'\n  <Override PartName="/themeVariants/variant{i}/theme/presentation.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        )
    files["[Content_Types].xml"] = _CONTENT_TYPES_TEMPLATE.format(
        extra_overrides=extra_ct
    )

    # Variant manager XML
    if manager_content is not None:
        files["themeVariants/themeVariantManager.xml"] = _VARIANT_MANAGER_TEMPLATE.format(
            content=manager_content
        )
    else:
        variants_xml = ""
        for i in range(1, variant_count + 1):
            variants_xml += (
                f'    <t:themeVariant name="Variant {i}" '
                f'vid="{{0000000{i}-0000-0000-0000-000000000000}}" '
                f'cx="12192000" cy="6858000" r:id="rId{i}"/>\n'
            )
        files["themeVariants/themeVariantManager.xml"] = _VARIANT_MANAGER_TEMPLATE.format(
            content=f"<t:themeVariantLst>\n{variants_xml}  </t:themeVariantLst>"
        )

    # Variant manager rels
    if manager_rels is not None:
        files["themeVariants/_rels/themeVariantManager.xml.rels"] = manager_rels
    else:
        rels_xml = ""
        for i in range(1, variant_count + 1):
            rels_xml += (
                f'  <Relationship Id="rId{i}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                f'Target="variant{i}/theme/presentation.xml"/>\n'
            )
        files["themeVariants/_rels/themeVariantManager.xml.rels"] = (
            _VARIANT_MANAGER_RELS_TEMPLATE.format(rels=rels_xml)
        )

    # Variant parts
    if include_variant_parts:
        for i in range(1, variant_count + 1):
            prefix = f"themeVariants/variant{i}/theme"
            files[f"{prefix}/presentation.xml"] = _PRESENTATION_XML
            files[f"{prefix}/_rels/presentation.xml.rels"] = (
                _VARIANT_PRESENTATION_RELS_TEMPLATE
            )
            files[f"{prefix}/theme/theme1.xml"] = variant_theme_xml

    return files


class TestSuperthemeValidation:
    """Tests for supertheme (themeVariantManager) validation."""

    def test_valid_supertheme(self, tmp_path: Path) -> None:
        """A well-formed supertheme should produce no supertheme-specific errors."""
        files = _base_pptx_files()
        files.update(_supertheme_files(variant_count=2))
        thmx = _build_thmx(tmp_path, "valid.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        # Filter to only supertheme-relevant errors
        st_errors = [
            e for e in result.errors
            if "themeVariant" in e.description or "variant" in e.description.lower()
        ]
        assert st_errors == [], [e.description for e in st_errors]

    def test_no_supertheme_no_errors(self, tmp_path: Path) -> None:
        """A regular PPTX without supertheme should have no supertheme errors."""
        files = _base_pptx_files()
        thmx = _build_thmx(tmp_path, "regular.pptx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        st_errors = [
            e for e in result.errors
            if "themeVariant" in e.description
        ]
        assert st_errors == []

    def test_missing_variant_list(self, tmp_path: Path) -> None:
        """themeVariantManager without themeVariantLst should error."""
        files = _base_pptx_files()
        files.update(
            _supertheme_files(
                variant_count=0,
                manager_content="<!-- empty -->",
            )
        )
        thmx = _build_thmx(tmp_path, "no_list.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("themeVariantLst" in d for d in descs), descs

    def test_empty_variant_list(self, tmp_path: Path) -> None:
        """themeVariantLst with no themeVariant children should error."""
        files = _base_pptx_files()
        files.update(
            _supertheme_files(
                variant_count=0,
                manager_content="<t:themeVariantLst/>",
            )
        )
        thmx = _build_thmx(tmp_path, "empty_list.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("at least one" in d for d in descs), descs

    def test_variant_missing_name(self, tmp_path: Path) -> None:
        """themeVariant without name attribute should error."""
        files = _base_pptx_files()
        # Build custom manager content with a variant missing name
        manager_content = (
            '<t:themeVariantLst>'
            '<t:themeVariant vid="{00000001-0000-0000-0000-000000000000}" '
            'cx="12192000" cy="6858000" r:id="rId1"/>'
            '</t:themeVariantLst>'
        )
        files.update(
            _supertheme_files(
                variant_count=1,
                manager_content=manager_content,
            )
        )
        thmx = _build_thmx(tmp_path, "no_name.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("'name'" in d for d in descs), descs

    def test_variant_invalid_guid(self, tmp_path: Path) -> None:
        """themeVariant with invalid vid (not a GUID) should error."""
        files = _base_pptx_files()
        manager_content = (
            '<t:themeVariantLst>'
            '<t:themeVariant name="Bad GUID" vid="not-a-guid" '
            'cx="12192000" cy="6858000" r:id="rId1"/>'
            '</t:themeVariantLst>'
        )
        files.update(
            _supertheme_files(
                variant_count=1,
                manager_content=manager_content,
            )
        )
        thmx = _build_thmx(tmp_path, "bad_guid.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("GUID" in d for d in descs), descs

    def test_variant_negative_dimensions(self, tmp_path: Path) -> None:
        """themeVariant with negative cx/cy should error."""
        files = _base_pptx_files()
        manager_content = (
            '<t:themeVariantLst>'
            '<t:themeVariant name="Bad Dims" '
            'vid="{00000001-0000-0000-0000-000000000000}" '
            'cx="-1" cy="6858000" r:id="rId1"/>'
            '</t:themeVariantLst>'
        )
        files.update(
            _supertheme_files(
                variant_count=1,
                manager_content=manager_content,
            )
        )
        thmx = _build_thmx(tmp_path, "bad_dims.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("positive" in d for d in descs), descs

    def test_variant_unresolved_relationship(self, tmp_path: Path) -> None:
        """themeVariant r:id pointing to nonexistent relationship should error."""
        files = _base_pptx_files()
        manager_content = (
            '<t:themeVariantLst>'
            '<t:themeVariant name="Orphan" '
            'vid="{00000001-0000-0000-0000-000000000000}" '
            'cx="12192000" cy="6858000" r:id="rId99"/>'
            '</t:themeVariantLst>'
        )
        # Empty rels (no rId99)
        empty_rels = _VARIANT_MANAGER_RELS_TEMPLATE.format(rels="")
        files.update(
            _supertheme_files(
                variant_count=0,
                manager_content=manager_content,
                manager_rels=empty_rels,
                include_variant_parts=False,
            )
        )
        thmx = _build_thmx(tmp_path, "orphan_rel.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("rId99" in d for d in descs), descs

    def test_variant_missing_target_part(self, tmp_path: Path) -> None:
        """Variant relationship target not present in package should error."""
        files = _base_pptx_files()
        files.update(
            _supertheme_files(
                variant_count=1,
                include_variant_parts=False,  # Don't create the actual parts
            )
        )
        thmx = _build_thmx(tmp_path, "missing_target.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("not found in package" in d for d in descs), descs

    def test_variant_theme_validated(self, tmp_path: Path) -> None:
        """Variant theme parts should be validated (e.g., missing clrScheme)."""
        bad_theme = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Bad">
  <a:themeElements>
    <a:fontScheme name="F">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Fmt">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""
        files = _base_pptx_files()
        files.update(
            _supertheme_files(
                variant_count=1,
                variant_theme_xml=bad_theme,
            )
        )
        thmx = _build_thmx(tmp_path, "bad_variant_theme.thmx", files)

        validator = OpenXmlValidator(schema_validation=False)
        result = validator.validate(thmx)
        descs = [e.description for e in result.errors]
        assert any("clrScheme" in d for d in descs), descs
