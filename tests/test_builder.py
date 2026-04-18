"""Unit tests for the OPC PackageBuilder."""

from __future__ import annotations

import zipfile
from io import BytesIO

from lxml import etree

from openxml_audit.builder import PackageBuilder
from openxml_audit.namespaces import CONTENT_TYPES, RELATIONSHIPS


def test_builder_produces_content_types_with_defaults_and_overrides() -> None:
    builder = PackageBuilder()
    builder.add_default_type("rels", "application/vnd.openxmlformats-package.relationships+xml")
    builder.add_default_type("xml", "application/xml")
    builder.add_part("/word/document.xml", b"<root/>", content_type="application/x-test")

    data = builder.to_bytes()

    with zipfile.ZipFile(BytesIO(data)) as zf:
        content_types = zf.read("[Content_Types].xml")
    root = etree.fromstring(content_types)
    assert root.tag == f"{{{CONTENT_TYPES}}}Types"
    defaults = root.findall(f"{{{CONTENT_TYPES}}}Default")
    overrides = root.findall(f"{{{CONTENT_TYPES}}}Override")
    assert {d.get("Extension") for d in defaults} == {"rels", "xml"}
    assert len(overrides) == 1
    assert overrides[0].get("PartName") == "/word/document.xml"
    assert overrides[0].get("ContentType") == "application/x-test"


def test_builder_emits_package_rels_at_canonical_path() -> None:
    builder = PackageBuilder()
    builder.add_relationship(
        "/", "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "word/document.xml",
    )

    data = builder.to_bytes()

    with zipfile.ZipFile(BytesIO(data)) as zf:
        assert "_rels/.rels" in zf.namelist()
        rels_xml = zf.read("_rels/.rels")
    root = etree.fromstring(rels_xml)
    assert root.tag == f"{{{RELATIONSHIPS}}}Relationships"
    rels = root.findall(f"{{{RELATIONSHIPS}}}Relationship")
    assert len(rels) == 1
    assert rels[0].get("Id") == "rId1"
    assert rels[0].get("Target") == "word/document.xml"


def test_builder_emits_part_rels_at_canonical_path() -> None:
    builder = PackageBuilder()
    builder.add_relationship(
        "/xl/workbook.xml", "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "worksheets/sheet1.xml",
    )

    data = builder.to_bytes()

    with zipfile.ZipFile(BytesIO(data)) as zf:
        assert "xl/_rels/workbook.xml.rels" in zf.namelist()


def test_builder_normalizes_relative_uris() -> None:
    builder = PackageBuilder()
    builder.add_part("word/document.xml", b"<root/>", content_type="application/x-test")
    data = builder.to_bytes()
    with zipfile.ZipFile(BytesIO(data)) as zf:
        assert "word/document.xml" in zf.namelist()
