"""Generated reference slides: pure-lxml slide authoring for Spec 034.

The index slide and the entrance-effect feature slides are authored at
build time rather than sourced from a committed scaffold. The timing
fragments reuse the same builders the oracle decks emit
(`oracle_starter_deck._build_anim_effect` et al.), so a generated
reference slide carries exactly the structure its capability finding
was registered from.
"""

from __future__ import annotations

from itertools import count
from typing import cast

from lxml import etree

from openxml_audit.namespaces import DRAWINGML, PRESENTATIONML
from openxml_audit.pptx.oracle_starter_deck import (
    BuildEntry,
    _build_anim_effect,
    _build_effect_par,
    _build_set_visibility,
)
from openxml_audit.pptx.timing_oracle_deck import _build_auto_timing_xml

__all__ = [
    "add_textbox",
    "build_entrance_fade_slide",
    "build_entrance_wipe_slide",
    "finish_slide",
    "new_slide",
]

_P = f"{{{PRESENTATIONML}}}"
_A = f"{{{DRAWINGML}}}"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_EMU_PER_INCH = 914400


def _emu(inches: float) -> str:
    return str(int(round(inches * _EMU_PER_INCH)))


def new_slide() -> tuple[etree._Element, etree._Element]:
    """Return (p:sld root, p:spTree) with the required scaffolding."""
    nsmap = {"a": DRAWINGML, "r": _R_NS, "p": PRESENTATIONML}
    sld = etree.Element(f"{_P}sld", nsmap=nsmap)
    c_sld = etree.SubElement(sld, f"{_P}cSld")
    sp_tree = etree.SubElement(c_sld, f"{_P}spTree")

    nv_grp = etree.SubElement(sp_tree, f"{_P}nvGrpSpPr")
    c_nv_pr = etree.SubElement(nv_grp, f"{_P}cNvPr")
    c_nv_pr.set("id", "1")
    c_nv_pr.set("name", "")
    etree.SubElement(nv_grp, f"{_P}cNvGrpSpPr")
    etree.SubElement(nv_grp, f"{_P}nvPr")
    grp_sp_pr = etree.SubElement(sp_tree, f"{_P}grpSpPr")
    xfrm = etree.SubElement(grp_sp_pr, f"{_A}xfrm")
    for tag in ("off", "ext", "chOff", "chExt"):
        el = etree.SubElement(xfrm, f"{_A}{tag}")
        if tag in ("off", "chOff"):
            el.set("x", "0")
            el.set("y", "0")
        else:
            el.set("cx", "0")
            el.set("cy", "0")
    return sld, sp_tree


def finish_slide(sld: etree._Element, *, timing_xml: str | None = None) -> bytes:
    """Append clrMapOvr (and optional authored timing) and serialize."""
    clr_map_ovr = etree.SubElement(sld, f"{_P}clrMapOvr")
    etree.SubElement(clr_map_ovr, f"{_A}masterClrMapping")
    if timing_xml is not None:
        sld.append(etree.fromstring(timing_xml.encode("utf-8")))
    return cast(
        bytes,
        etree.tostring(sld, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _shape_properties(
    sp: etree._Element,
    *,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    fill_hex: str | None,
) -> None:
    sp_pr = etree.SubElement(sp, f"{_P}spPr")
    xfrm = etree.SubElement(sp_pr, f"{_A}xfrm")
    off = etree.SubElement(xfrm, f"{_A}off")
    off.set("x", _emu(left_in))
    off.set("y", _emu(top_in))
    ext = etree.SubElement(xfrm, f"{_A}ext")
    ext.set("cx", _emu(width_in))
    ext.set("cy", _emu(height_in))
    prst_geom = etree.SubElement(sp_pr, f"{_A}prstGeom")
    prst_geom.set("prst", "rect")
    etree.SubElement(prst_geom, f"{_A}avLst")
    if fill_hex is not None:
        solid_fill = etree.SubElement(sp_pr, f"{_A}solidFill")
        srgb = etree.SubElement(solid_fill, f"{_A}srgbClr")
        srgb.set("val", fill_hex)


def _text_body(
    sp: etree._Element,
    *,
    text: str,
    size: int,
    bold: bool,
    center: bool,
) -> None:
    tx_body = etree.SubElement(sp, f"{_P}txBody")
    body_pr = etree.SubElement(tx_body, f"{_A}bodyPr")
    body_pr.set("wrap", "square")
    etree.SubElement(tx_body, f"{_A}lstStyle")
    paragraph = etree.SubElement(tx_body, f"{_A}p")
    if center:
        p_pr = etree.SubElement(paragraph, f"{_A}pPr")
        p_pr.set("algn", "ctr")
    run = etree.SubElement(paragraph, f"{_A}r")
    r_pr = etree.SubElement(run, f"{_A}rPr")
    r_pr.set("lang", "en-US")
    r_pr.set("sz", str(size))
    if bold:
        r_pr.set("b", "1")
    t = etree.SubElement(run, f"{_A}t")
    t.text = text


def add_textbox(
    sp_tree: etree._Element,
    shape_id: int,
    *,
    top_in: float,
    text: str,
    size: int,
    bold: bool,
    left_in: float = 0.55,
    width_in: float = 12.2,
    height_in: float = 0.4,
) -> int:
    """Add a plain textbox shape; returns the next free shape id."""
    sp = etree.SubElement(sp_tree, f"{_P}sp")
    nv_sp_pr = etree.SubElement(sp, f"{_P}nvSpPr")
    c_nv_pr = etree.SubElement(nv_sp_pr, f"{_P}cNvPr")
    c_nv_pr.set("id", str(shape_id))
    c_nv_pr.set("name", f"Reference TextBox {shape_id}")
    c_nv_sp_pr = etree.SubElement(nv_sp_pr, f"{_P}cNvSpPr")
    c_nv_sp_pr.set("txBox", "1")
    etree.SubElement(nv_sp_pr, f"{_P}nvPr")
    _shape_properties(
        sp,
        left_in=left_in,
        top_in=top_in,
        width_in=width_in,
        height_in=height_in,
        fill_hex=None,
    )
    _text_body(sp, text=text, size=size, bold=bold, center=False)
    return shape_id + 1


def _add_target_box(
    sp_tree: etree._Element,
    shape_id: int,
    *,
    label: str,
    fill_hex: str,
) -> int:
    """Add the filled rectangle an entrance effect targets."""
    sp = etree.SubElement(sp_tree, f"{_P}sp")
    nv_sp_pr = etree.SubElement(sp, f"{_P}nvSpPr")
    c_nv_pr = etree.SubElement(nv_sp_pr, f"{_P}cNvPr")
    c_nv_pr.set("id", str(shape_id))
    c_nv_pr.set("name", f"Target {shape_id}")
    etree.SubElement(nv_sp_pr, f"{_P}cNvSpPr")
    etree.SubElement(nv_sp_pr, f"{_P}nvPr")
    _shape_properties(
        sp,
        left_in=4.9,
        top_in=2.9,
        width_in=3.5,
        height_in=1.6,
        fill_hex=fill_hex,
    )
    _text_body(sp, text=label, size=1400, bold=True, center=True)
    return shape_id + 1


def _build_entrance_slide(
    *,
    title: str,
    subtitle: str,
    target_label: str,
    filter_name: str,
    preset_id: int,
    preset_subtype: int,
) -> bytes:
    sld, sp_tree = new_slide()
    shape_id = 2
    shape_id = add_textbox(sp_tree, shape_id, top_in=0.4, text=title, size=2000, bold=True)
    shape_id = add_textbox(sp_tree, shape_id, top_in=1.0, text=subtitle, size=1200, bold=False)
    target_id = shape_id
    shape_id = _add_target_box(sp_tree, target_id, label=target_label, fill_hex="93C5FD")

    id_counter = count(shape_id + 10)
    effect = _build_effect_par(
        id_counter=id_counter,
        duration_ms=500,
        node_type="clickEffect",
        preset_id=preset_id,
        preset_class="entr",
        preset_subtype=preset_subtype,
        grp_id=1,
        child_elements=[
            _build_set_visibility(
                id_counter=id_counter,
                target_shape_id=target_id,
                visibility="visible",
            ),
            _build_anim_effect(
                id_counter=id_counter,
                target_shape_id=target_id,
                duration_ms=500,
                transition="in",
                filter_name=filter_name,
            ),
        ],
    )
    timing_xml = _build_auto_timing_xml(
        start_id=shape_id + 1,
        main_effect_pars=(effect,),
        build_entries=(BuildEntry(shape_id=target_id, grp_id=1),),
    )
    return finish_slide(sld, timing_xml=timing_xml)


def build_entrance_fade_slide() -> bytes:
    """pptx.anim.effect.entr.fade: native animEffect transition=in filter=fade."""
    return _build_entrance_slide(
        title="Entrance Fade",
        subtitle=(
            "On slide entry the target fades in: <p:animEffect transition='in' filter='fade'>."
        ),
        target_label="fade target",
        filter_name="fade",
        preset_id=10,
        preset_subtype=0,
    )


def build_entrance_wipe_slide() -> bytes:
    """pptx.anim.effect.entr.wipe: native animEffect transition=in filter=wipe(...)."""
    return _build_entrance_slide(
        title="Entrance Wipe",
        subtitle=(
            "On slide entry the target wipes in from the bottom: "
            "<p:animEffect transition='in' filter='wipe(up)'>. Direction/"
            "subtype semantics are narrower than base wipe support "
            "(see the finding's constraints)."
        ),
        target_label="wipe target",
        filter_name="wipe(up)",
        preset_id=22,
        preset_subtype=1,
    )
