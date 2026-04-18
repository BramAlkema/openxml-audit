from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree as ET

from openxml_audit.pptx.oracle_deck_scaffold import (
    inject_timing_map_into_pptx,
    patch_slide_xml_with_timing,
    save_oracle_presentation,
)

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def test_patch_slide_xml_with_timing_replaces_existing_node() -> None:
    slide_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld/>
  <p:timing><p:tnLst/></p:timing>
</p:sld>"""
    timing_xml = (
        '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:tnLst><p:par/></p:tnLst></p:timing>'
    )

    patched = patch_slide_xml_with_timing(slide_xml, timing_xml)
    root = ET.fromstring(patched)

    assert root.xpath("count(.//p:timing)", namespaces=NS) == 1.0
    assert root.xpath("count(.//p:timing/p:tnLst/p:par)", namespaces=NS) == 1.0


def _write_minimal_pptx(path: Path) -> None:
    slide_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld/>
</p:sld>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ppt/slides/slide1.xml", slide_xml)


def test_inject_timing_map_into_pptx_patches_target_slide(tmp_path: Path) -> None:
    pptx_path = tmp_path / "deck.pptx"
    _write_minimal_pptx(pptx_path)
    timing_xml = (
        '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:tnLst><p:par/></p:tnLst></p:timing>'
    )

    inject_timing_map_into_pptx(pptx_path, {1: timing_xml})

    with zipfile.ZipFile(pptx_path) as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))

    assert root.xpath("count(.//p:timing/p:tnLst/p:par)", namespaces=NS) == 1.0


class _StubPresentation:
    def __init__(self, slide_xml: str) -> None:
        self._slide_xml = slide_xml

    def save(self, output_path: Path) -> None:
        with zipfile.ZipFile(output_path, "w") as archive:
            archive.writestr("ppt/slides/slide1.xml", self._slide_xml)


def test_save_oracle_presentation_uses_scaffold_then_patches_timing(tmp_path: Path) -> None:
    presentation = _StubPresentation(
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld/>
</p:sld>"""
    )
    output_path = tmp_path / "oracle.pptx"
    timing_xml = (
        '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:tnLst><p:par/></p:tnLst></p:timing>'
    )

    deck_path = save_oracle_presentation(
        presentation,
        output_path,
        timing_by_slide_number={1: timing_xml},
        temp_prefix="ppt-test-scaffold-",
    )

    assert deck_path == output_path
    with zipfile.ZipFile(deck_path) as archive:
        root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))

    assert root.xpath("count(.//p:timing/p:tnLst/p:par)", namespaces=NS) == 1.0
