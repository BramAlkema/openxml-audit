"""Tests for cross-part semantic constraints."""

from __future__ import annotations

import zipfile
from pathlib import Path

from openxml_audit.context import ValidationContext
from openxml_audit.package import OpenXmlPackage
from openxml_audit.parts import OpenXmlPart
from openxml_audit.semantic.constraints import CrossPartCountConstraint
from tests.fixture_loader import load_fixture_bytes


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSMAP = {"x": SPREADSHEET_NS}


def _write_zip(path: Path, parts: dict[str, bytes | str]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, payload in parts.items():
            data = payload.encode("utf-8") if isinstance(payload, str) else payload
            zf.writestr(name.lstrip("/"), data)


def test_cross_part_count_direct_part_path(tmp_path: Path) -> None:
    target_xml = load_fixture_bytes("cross_part", "target.xml")
    current_xml = load_fixture_bytes("cross_part", "current.xml")

    zip_path = tmp_path / "cross-part.pptx"
    _write_zip(zip_path, {"current.xml": current_xml, "target.xml": target_xml})

    with OpenXmlPackage(zip_path) as package:
        context = ValidationContext(package=package)
        part = OpenXmlPart(package, "/current.xml")
        context.set_part(part)

        element = part.xml
        assert element is not None

        constraint = CrossPartCountConstraint(
            attribute="ref",
            part_path="/target.xml",
            element_xpath="x:item",
            count_offset=1,
            namespace_map=NSMAP,
        )

        assert constraint.validate(element, context)

        element.set("ref", "3")
        context.clear_errors()
        assert not constraint.validate(element, context)


def test_cross_part_count_current_part(tmp_path: Path) -> None:
    current_xml = load_fixture_bytes("cross_part", "current_with_items.xml")

    zip_path = tmp_path / "current-part.pptx"
    _write_zip(zip_path, {"current.xml": current_xml})

    with OpenXmlPackage(zip_path) as package:
        context = ValidationContext(package=package)
        part = OpenXmlPart(package, "/current.xml")
        context.set_part(part)

        element = part.xml
        assert element is not None

        constraint = CrossPartCountConstraint(
            attribute="ref",
            part_path=".",
            element_xpath="x:item",
            count_offset=0,
            namespace_map=NSMAP,
        )

        assert constraint.validate(element, context)

        element.set("ref", "2")
        context.clear_errors()
        assert not constraint.validate(element, context)
