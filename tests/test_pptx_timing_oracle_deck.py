from __future__ import annotations

import zipfile

from lxml import etree as ET

from openxml_audit.pptx.timing_oracle_deck import build_timing_oracle_deck

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def _slide_root(pptx_path, slide_number: int) -> ET._Element:
    with zipfile.ZipFile(pptx_path) as pptx:
        slide_xml = pptx.read(f"ppt/slides/slide{slide_number}.xml")
    return ET.fromstring(slide_xml)


def test_timing_oracle_deck_contains_expected_native_timing_fields(tmp_path) -> None:
    deck_path = build_timing_oracle_deck(tmp_path / "timing-oracle.pptx")

    assert deck_path.exists()

    slide2 = _slide_root(deck_path, 2)
    slide3 = _slide_root(deck_path, 3)
    slide4 = _slide_root(deck_path, 4)
    slide5 = _slide_root(deck_path, 5)
    slide6 = _slide_root(deck_path, 6)

    for slide in (slide2, slide3, slide4, slide5, slide6):
        assert (
            slide.xpath(
                "count(.//p:cTn[@nodeType='interactiveSeq'])",
                namespaces=NS,
            )
            == 0.0
        )

    assert (
        slide2.xpath(
            "count(.//p:endCondLst/p:cond[@delay='2000'])",
            namespaces=NS,
        )
        == 1.0
    )
    assert (
        slide2.xpath(
            "count(.//p:seq/p:cTn[@nodeType='mainSeq']/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond[@delay='0'])",
            namespaces=NS,
        )
        >= 1.0
    )
    assert (
        slide3.xpath(
            "count(.//p:endCondLst/p:cond[@evt='onClick'])",
            namespaces=NS,
        )
        == 1.0
    )
    assert (
        slide4.xpath(
            "count(.//p:cTn[@repeatCount='indefinite'][@repeatDur='2500'])",
            namespaces=NS,
        )
        == 1.0
    )
    assert (
        slide5.xpath(
            "count(.//p:cTn[@nodeType='clickEffect'][@restart='always'])",
            namespaces=NS,
        )
        == 1.0
    )
    assert (
        slide5.xpath(
            "count(.//p:cTn[@nodeType='clickEffect'][@restart='whenNotActive'])",
            namespaces=NS,
        )
        == 1.0
    )
    assert (
        slide5.xpath(
            "count(.//p:cTn[@nodeType='clickEffect'][@restart='never'])",
            namespaces=NS,
        )
        == 1.0
    )
    assert (
        slide6.xpath(
            "count(.//p:cTn[@nodeType='clickEffect'][@restart='never']/p:stCondLst/p:cond[@delay='5000'])",
            namespaces=NS,
        )
        == 1.0
    )
