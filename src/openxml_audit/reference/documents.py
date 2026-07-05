"""Per-format canonical reference document builders (Spec 034).

Each builder assembles one reference document from the capability
ledger, writes a provenance manifest next to it, and self-validates the
result with `OpenXmlValidator` — a reference document that fails our own
tier-1 floor is a build error, not an artifact.

Builds are deterministic: no timestamps, ledger-sorted ordering, fixed
zip metadata. Same inputs, same bytes.
"""

from __future__ import annotations

import json
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, cast

from lxml import etree

from openxml_audit import __version__
from openxml_audit.builder import PackageBuilder
from openxml_audit.evidence import EvidenceTier
from openxml_audit.namespaces import (
    CONTENT_TYPES,
    DRAWINGML,
    PRESENTATIONML,
    REL_FONT_TABLE,
    REL_OFFICE_DOCUMENT,
    REL_SETTINGS,
    REL_SHARED_STRINGS,
    REL_SLIDE,
    REL_STYLES,
    REL_THEME,
    RELATIONSHIPS,
    SPREADSHEETML,
    WORDPROCESSINGML,
)
from openxml_audit.pptx.oracle_deck_scaffold import scaffold_root
from openxml_audit.reference.emitters import (
    DOCX_BODY_EMITTERS,
    PPTX_SLIDE_SOURCES,
    XLSX_ROW_EMITTERS,
    has_emitter,
)
from openxml_audit.reference.ledger import (
    LedgerEntry,
    collect_ledger,
    qualifies_at,
)
from openxml_audit.validator import OpenXmlValidator

__all__ = [
    "ReferenceBuildError",
    "ReferenceBuildResult",
    "build_reference_document",
]

_P = f"{{{PRESENTATIONML}}}"
_A = f"{{{DRAWINGML}}}"
_W = f"{{{WORDPROCESSINGML}}}"
_S = f"{{{SPREADSHEETML}}}"
_CT = f"{{{CONTENT_TYPES}}}"
_PKG_R = f"{{{RELATIONSHIPS}}}"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

_EXTENSIONS = {"docx": ".docx", "pptx": ".pptx", "xlsx": ".xlsx"}

# Fixed zip timestamp: reference documents are derived artifacts and
# must be byte-reproducible run-to-run.
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)

_EMU_PER_INCH = 914400


class ReferenceBuildError(Exception):
    """Raised when a reference document cannot be built or fails self-validation."""


@dataclass(frozen=True, slots=True)
class ReferenceBuildResult:
    """Outcome of one per-format reference build."""

    format: str
    document_path: Path
    manifest_path: Path
    included_keys: tuple[str, ...]
    excluded: tuple[tuple[str, str], ...]  # (key, reason)
    manifest: dict[str, Any] = field(repr=False)


def build_reference_document(
    fmt: str,
    output_dir: Path | str,
    *,
    minimum_tier: EvidenceTier = EvidenceTier.ROUNDTRIP_PRESERVED,
) -> ReferenceBuildResult:
    """Build the canonical reference document for one format.

    Returns the build result after the document passes self-validation.
    """
    if fmt not in _EXTENSIONS:
        raise ReferenceBuildError(f"Unknown reference format: {fmt!r}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    document_path = output_dir / f"reference{_EXTENSIONS[fmt]}"
    manifest_path = output_dir / f"reference{_EXTENSIONS[fmt]}.manifest.json"

    entries = [entry for entry in collect_ledger() if entry.format == fmt]
    included, excluded = _partition(fmt, entries, minimum_tier)

    if fmt == "pptx":
        _build_pptx(document_path, included)
    elif fmt == "docx":
        _build_docx(document_path, included)
    else:
        _build_xlsx(document_path, included)

    _self_validate(document_path)

    manifest = _build_manifest(fmt, minimum_tier, included, excluded)
    manifest_path.write_text(
        json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )

    return ReferenceBuildResult(
        format=fmt,
        document_path=document_path,
        manifest_path=manifest_path,
        included_keys=tuple(entry.finding.key for entry in included),
        excluded=tuple((entry.finding.key, reason) for entry, reason in excluded),
        manifest=manifest,
    )


def _partition(
    fmt: str,
    entries: list[LedgerEntry],
    minimum_tier: EvidenceTier,
) -> tuple[list[LedgerEntry], list[tuple[LedgerEntry, str]]]:
    included: list[LedgerEntry] = []
    excluded: list[tuple[LedgerEntry, str]] = []
    for entry in entries:
        if not qualifies_at(entry.finding, minimum_tier):
            excluded.append((entry, "below-minimum-tier"))
        elif not has_emitter(fmt, entry.finding.key):
            excluded.append((entry, "no-emitter"))
        else:
            included.append(entry)
    return included, excluded


def _self_validate(document_path: Path) -> None:
    result = OpenXmlValidator().validate(document_path)
    if result.error_count > 0:
        details = "; ".join(
            f"{error.error_type}: {error.description}" for error in result.errors[:5]
        )
        raise ReferenceBuildError(
            f"Generated reference document failed self-validation "
            f"({result.error_count} errors): {details}"
        )


def _build_manifest(
    fmt: str,
    minimum_tier: EvidenceTier,
    included: list[LedgerEntry],
    excluded: list[tuple[LedgerEntry, str]],
) -> dict[str, Any]:
    features = []
    for position, entry in enumerate(included):
        payload = entry.finding.to_dict()
        payload["included"] = True
        payload["location"] = _location(fmt, position)
        features.append(payload)
    excluded_payload = []
    for entry, reason in excluded:
        payload = entry.finding.to_dict()
        payload["included"] = False
        payload["reason"] = reason
        excluded_payload.append(payload)
    return {
        "format": fmt,
        "spec": "034-canonical-reference-documents",
        "generator": f"openxml-audit {__version__} reference builder",
        "minimum_tier": minimum_tier.value,
        "features": features,
        "excluded": excluded_payload,
    }


def _location(fmt: str, position: int) -> str:
    if fmt == "pptx":
        # Slide 1 is the generated index; features may span several
        # slides, so the manifest points at the first one and the index
        # slide carries the full map.
        return f"slide {position + 2}"
    if fmt == "docx":
        return f"section {position + 1}"
    return f"row {position + 3}"


# --- deterministic zip ------------------------------------------------------


def _write_deterministic_zip(path: Path, members: list[tuple[str, bytes]]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for name, payload in sorted(members):
            info = zipfile.ZipInfo(name, date_time=_ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, payload)


def _normalize_zip(path: Path) -> None:
    """Rewrite a zip in place with deterministic member metadata."""
    with zipfile.ZipFile(path, "r") as zf:
        members = [(info.filename, zf.read(info.filename)) for info in zf.infolist()]
    _write_deterministic_zip(path, members)


# --- PPTX -------------------------------------------------------------------


def _build_pptx(document_path: Path, included: list[LedgerEntry]) -> None:
    root = scaffold_root("timing_oracle")

    # Feature slides in ledger order, sourced from committed scaffolds.
    slide_sources: list[tuple[LedgerEntry, Path]] = []
    for entry in included:
        for source in PPTX_SLIDE_SOURCES[entry.finding.key]:
            source_root = (
                root if source.scaffold == "timing_oracle" else scaffold_root(source.scaffold)
            )
            slide_sources.append(
                (entry, source_root / "ppt" / "slides" / f"slide{source.slide_number}.xml")
            )

    members: list[tuple[str, bytes]] = []
    for file_path in sorted(p for p in root.rglob("*") if p.is_file()):
        arcname = file_path.relative_to(root).as_posix()
        if arcname.startswith("ppt/slides/"):
            continue
        if arcname in {
            "[Content_Types].xml",
            "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels",
        }:
            continue
        members.append((arcname, file_path.read_bytes()))

    # Slide 1 is the generated index; feature slides follow, renumbered.
    slide_payloads: list[tuple[bytes, bytes]] = [
        (_build_index_slide_xml(included), _index_slide_rels(root))
    ]
    for _entry, source_path in slide_sources:
        rels_path = source_path.parent / "_rels" / f"{source_path.name}.rels"
        slide_payloads.append((source_path.read_bytes(), rels_path.read_bytes()))

    for number, (slide_xml, rels_xml) in enumerate(slide_payloads, start=1):
        members.append((f"ppt/slides/slide{number}.xml", slide_xml))
        members.append((f"ppt/slides/_rels/slide{number}.xml.rels", rels_xml))

    slide_count = len(slide_payloads)
    rels_payload, slide_rids = _rewrite_presentation_rels(root, slide_count)
    members.append(("ppt/presentation.xml", _rewrite_presentation_xml(root, slide_rids)))
    members.append(("ppt/_rels/presentation.xml.rels", rels_payload))
    members.append(("[Content_Types].xml", _rewrite_content_types(root, slide_count)))

    _write_deterministic_zip(document_path, members)


def _rewrite_presentation_rels(root: Path, slide_count: int) -> tuple[bytes, list[str]]:
    """Drop scaffold slide rels; append fresh sequential slide rels.

    Returns the serialized rels and the new slide rIds in slide order.
    """
    rels_root = etree.fromstring((root / "ppt/_rels/presentation.xml.rels").read_bytes())
    kept_ids: list[int] = []
    for rel in list(rels_root):
        if rel.get("Type") == REL_SLIDE:
            rels_root.remove(rel)
        else:
            rel_id = rel.get("Id", "")
            if rel_id.startswith("rId") and rel_id[3:].isdigit():
                kept_ids.append(int(rel_id[3:]))
    next_id = max(kept_ids, default=0) + 1
    slide_rids: list[str] = []
    for number in range(1, slide_count + 1):
        rid = f"rId{next_id}"
        next_id += 1
        slide_rids.append(rid)
        rel = etree.SubElement(rels_root, f"{_PKG_R}Relationship")
        rel.set("Id", rid)
        rel.set("Type", REL_SLIDE)
        rel.set("Target", f"slides/slide{number}.xml")
    payload = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return payload, slide_rids


def _rewrite_presentation_xml(root: Path, slide_rids: list[str]) -> bytes:
    pres_root = etree.fromstring((root / "ppt/presentation.xml").read_bytes())
    sld_id_lst = pres_root.find(f"{_P}sldIdLst")
    if sld_id_lst is None:
        raise ReferenceBuildError("Scaffold presentation.xml is missing <p:sldIdLst>")
    for child in list(sld_id_lst):
        sld_id_lst.remove(child)
    for index, rid in enumerate(slide_rids):
        sld_id = etree.SubElement(sld_id_lst, f"{_P}sldId")
        sld_id.set("id", str(256 + index))
        sld_id.set(f"{{{_R_NS}}}id", rid)
    return cast(
        bytes,
        etree.tostring(pres_root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _rewrite_content_types(root: Path, slide_count: int) -> bytes:
    types_root = etree.fromstring((root / "[Content_Types].xml").read_bytes())
    for override in list(types_root.findall(f"{_CT}Override")):
        part_name = override.get("PartName", "")
        if part_name.startswith("/ppt/slides/"):
            types_root.remove(override)
    for number in range(1, slide_count + 1):
        override = etree.SubElement(types_root, f"{_CT}Override")
        override.set("PartName", f"/ppt/slides/slide{number}.xml")
        override.set("ContentType", _CT_SLIDE)
    return cast(
        bytes,
        etree.tostring(types_root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _index_slide_rels(root: Path) -> bytes:
    # The index slide uses the same layout as the scaffold's own slides.
    return (root / "ppt/slides/_rels/slide1.xml.rels").read_bytes()


def _emu(inches: float) -> str:
    return str(int(round(inches * _EMU_PER_INCH)))


def _build_index_slide_xml(included: list[LedgerEntry]) -> bytes:
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

    shape_id = 2
    shape_id = _add_index_textbox(
        sp_tree,
        shape_id,
        top_in=0.4,
        text="openxml-audit Canonical Reference (PPTX)",
        size=2400,
        bold=True,
    )
    shape_id = _add_index_textbox(
        sp_tree,
        shape_id,
        top_in=1.0,
        text=(
            "Generated from the capability ledger (Spec 034). "
            "Each following slide carries one proven feature."
        ),
        size=1200,
        bold=False,
    )
    top = 1.7
    slide_number = 2
    for entry in included:
        tiers = ", ".join(tier.value for tier in entry.finding.evidence_tiers)
        span = len(PPTX_SLIDE_SOURCES[entry.finding.key])
        slides = (
            f"slide {slide_number}"
            if span == 1
            else f"slides {slide_number}-{slide_number + span - 1}"
        )
        shape_id = _add_index_textbox(
            sp_tree,
            shape_id,
            top_in=top,
            text=f"{slides}: {entry.finding.key} [{tiers}]",
            size=1300,
            bold=False,
        )
        top += 0.45
        slide_number += span
    if not included:
        _add_index_textbox(
            sp_tree,
            shape_id,
            top_in=top,
            text="No capability findings qualify at the requested tier yet.",
            size=1300,
            bold=False,
        )

    clr_map_ovr = etree.SubElement(sld, f"{_P}clrMapOvr")
    etree.SubElement(clr_map_ovr, f"{_A}masterClrMapping")
    return cast(
        bytes,
        etree.tostring(sld, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _add_index_textbox(
    sp_tree: etree._Element,
    shape_id: int,
    *,
    top_in: float,
    text: str,
    size: int,
    bold: bool,
) -> int:
    sp = etree.SubElement(sp_tree, f"{_P}sp")
    nv_sp_pr = etree.SubElement(sp, f"{_P}nvSpPr")
    c_nv_pr = etree.SubElement(nv_sp_pr, f"{_P}cNvPr")
    c_nv_pr.set("id", str(shape_id))
    c_nv_pr.set("name", f"Index TextBox {shape_id}")
    c_nv_sp_pr = etree.SubElement(nv_sp_pr, f"{_P}cNvSpPr")
    c_nv_sp_pr.set("txBox", "1")
    etree.SubElement(nv_sp_pr, f"{_P}nvPr")

    sp_pr = etree.SubElement(sp, f"{_P}spPr")
    xfrm = etree.SubElement(sp_pr, f"{_A}xfrm")
    off = etree.SubElement(xfrm, f"{_A}off")
    off.set("x", _emu(0.55))
    off.set("y", _emu(top_in))
    ext = etree.SubElement(xfrm, f"{_A}ext")
    ext.set("cx", _emu(12.2))
    ext.set("cy", _emu(0.4))
    prst_geom = etree.SubElement(sp_pr, f"{_A}prstGeom")
    prst_geom.set("prst", "rect")
    etree.SubElement(prst_geom, f"{_A}avLst")

    tx_body = etree.SubElement(sp, f"{_P}txBody")
    body_pr = etree.SubElement(tx_body, f"{_A}bodyPr")
    body_pr.set("wrap", "square")
    etree.SubElement(tx_body, f"{_A}lstStyle")
    paragraph = etree.SubElement(tx_body, f"{_A}p")
    run = etree.SubElement(paragraph, f"{_A}r")
    r_pr = etree.SubElement(run, f"{_A}rPr")
    r_pr.set("lang", "en-US")
    r_pr.set("sz", str(size))
    if bold:
        r_pr.set("b", "1")
    t = etree.SubElement(run, f"{_A}t")
    t.text = text
    return shape_id + 1


# --- DOCX -------------------------------------------------------------------


def _build_docx(document_path: Path, included: list[LedgerEntry]) -> None:
    nsmap = {"w": WORDPROCESSINGML}
    body = etree.Element(f"{_W}body", nsmap=nsmap)
    _add_docx_paragraph(
        body, "openxml-audit Canonical Reference (DOCX)", bold=True, size_half_points=32
    )
    _add_docx_paragraph(
        body,
        "Generated from the capability ledger (Spec 034). "
        "Each section below carries one proven feature.",
    )
    for entry in included:
        tiers = ", ".join(tier.value for tier in entry.finding.evidence_tiers)
        _add_docx_paragraph(body, f"{entry.finding.key} [{tiers}]", bold=True, size_half_points=26)
        _add_docx_paragraph(body, entry.finding.summary)
        for block in DOCX_BODY_EMITTERS[entry.finding.key]():
            body.append(block)
    if not included:
        _add_docx_paragraph(body, "No capability findings qualify at the requested tier yet.")
    body.append(_docx_sect_pr())
    _write_docx_package(document_path, body)
    _normalize_zip(document_path)


_CT_DOCX_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
_CT_DOCX_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_CT_DOCX_SETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
_CT_DOCX_FONT_TABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
_CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
_CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"

_DOCX_STYLES_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b"<w:docDefaults><w:rPrDefault><w:rPr>"
    b'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
    b'<w:sz w:val="22"/>'
    b"</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
    b"</w:styles>"
)
_DOCX_SETTINGS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)
_DOCX_FONT_TABLE_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b'<w:font w:name="Calibri"><w:pitch w:val="variable"/></w:font>'
    b"</w:fonts>"
)


def _scaffold_theme_xml() -> bytes:
    """Reuse the committed scaffold's Office theme (format-neutral DrawingML)."""
    return (scaffold_root("timing_oracle") / "ppt/theme/theme1.xml").read_bytes()


def _write_docx_package(document_path: Path, body: etree._Element) -> None:
    """Assemble the reference DOCX with the support parts Word requires.

    Unlike the calibration-probe `build_minimal_docx`, the reference
    document must clear the validator's app-compat relationship checks
    (styles, settings, fontTable, theme).
    """
    document = etree.Element(f"{_W}document", nsmap={"w": WORDPROCESSINGML})
    document.append(body)
    document_xml = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)

    builder = PackageBuilder()
    builder.add_default_type("rels", _CT_RELS)
    builder.add_default_type("xml", "application/xml")
    builder.add_relationship("/", "rId1", REL_OFFICE_DOCUMENT, "word/document.xml")
    builder.add_part("/word/document.xml", document_xml, content_type=_CT_DOCX_MAIN)
    for rid, rel_type, target, part, xml, content_type in (
        ("rId1", REL_STYLES, "styles.xml", "/word/styles.xml", _DOCX_STYLES_XML, _CT_DOCX_STYLES),
        (
            "rId2",
            REL_SETTINGS,
            "settings.xml",
            "/word/settings.xml",
            _DOCX_SETTINGS_XML,
            _CT_DOCX_SETTINGS,
        ),
        (
            "rId3",
            REL_FONT_TABLE,
            "fontTable.xml",
            "/word/fontTable.xml",
            _DOCX_FONT_TABLE_XML,
            _CT_DOCX_FONT_TABLE,
        ),
        (
            "rId4",
            REL_THEME,
            "theme/theme1.xml",
            "/word/theme/theme1.xml",
            _scaffold_theme_xml(),
            _CT_THEME,
        ),
    ):
        builder.add_relationship("/word/document.xml", rid, rel_type, target)
        builder.add_part(part, xml, content_type=content_type)
    builder.write(document_path)


def _add_docx_paragraph(
    body: etree._Element,
    text: str,
    *,
    bold: bool = False,
    size_half_points: int | None = None,
) -> None:
    paragraph = etree.SubElement(body, f"{_W}p")
    run = etree.SubElement(paragraph, f"{_W}r")
    if bold or size_half_points is not None:
        r_pr = etree.SubElement(run, f"{_W}rPr")
        if bold:
            etree.SubElement(r_pr, f"{_W}b")
        if size_half_points is not None:
            sz = etree.SubElement(r_pr, f"{_W}sz")
            sz.set(f"{_W}val", str(size_half_points))
    t = etree.SubElement(run, f"{_W}t")
    t.text = text


def _docx_sect_pr() -> etree._Element:
    sect_pr = etree.Element(f"{_W}sectPr")
    pg_sz = etree.SubElement(sect_pr, f"{_W}pgSz")
    pg_sz.set(f"{_W}w", "12240")
    pg_sz.set(f"{_W}h", "15840")
    pg_mar = etree.SubElement(sect_pr, f"{_W}pgMar")
    for attr, value in (
        ("top", "1440"),
        ("right", "1440"),
        ("bottom", "1440"),
        ("left", "1440"),
        ("header", "720"),
        ("footer", "720"),
        ("gutter", "0"),
    ):
        pg_mar.set(f"{_W}{attr}", value)
    cols = etree.SubElement(sect_pr, f"{_W}cols")
    cols.set(f"{_W}space", "720")
    doc_grid = etree.SubElement(sect_pr, f"{_W}docGrid")
    doc_grid.set(f"{_W}linePitch", "360")
    return sect_pr


# --- XLSX -------------------------------------------------------------------


_CT_XLSX_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
_CT_XLSX_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_CT_XLSX_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
_CT_XLSX_SHARED_STRINGS = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)
_REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

_XLSX_STYLES_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    b'<fills count="2"><fill><patternFill patternType="none"/></fill>'
    b'<fill><patternFill patternType="gray125"/></fill></fills>'
    b'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    b'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    b'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    b'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    b"</styleSheet>"
)


class _SharedStrings:
    """Shared-string interner: text -> stable SST index."""

    def __init__(self) -> None:
        self._indexes: dict[str, int] = {}
        self._uses = 0

    def intern(self, text: str) -> int:
        self._uses += 1
        return self._indexes.setdefault(text, len(self._indexes))

    def to_xml(self) -> bytes:
        sst = etree.Element(f"{_S}sst", nsmap={None: SPREADSHEETML})
        sst.set("count", str(self._uses))
        sst.set("uniqueCount", str(len(self._indexes)))
        for text in self._indexes:
            si = etree.SubElement(sst, f"{_S}si")
            t = etree.SubElement(si, f"{_S}t")
            t.text = text
        return cast(
            bytes,
            etree.tostring(sst, xml_declaration=True, encoding="UTF-8", standalone=True),
        )


def _build_xlsx(document_path: Path, included: list[LedgerEntry]) -> None:
    strings = _SharedStrings()
    sheet_data = etree.Element(f"{_S}sheetData", nsmap={None: SPREADSHEETML})
    _add_xlsx_row(sheet_data, strings, 1, ["openxml-audit Canonical Reference (XLSX) — Spec 034"])
    _add_xlsx_row(sheet_data, strings, 2, ["feature", "summary", "evidence tiers"])
    row_number = 3
    for entry in included:
        tiers = ", ".join(tier.value for tier in entry.finding.evidence_tiers)
        _add_xlsx_row(
            sheet_data,
            strings,
            row_number,
            [entry.finding.key, entry.finding.summary, tiers],
        )
        row_number += 1
        for row in XLSX_ROW_EMITTERS[entry.finding.key](strings.intern):
            row.set("r", str(row_number))
            sheet_data.append(row)
            row_number += 1
    if not included:
        _add_xlsx_row(
            sheet_data,
            strings,
            row_number,
            ["(none)", "No capability findings qualify at the requested tier yet.", ""],
        )

    worksheet = etree.Element(f"{_S}worksheet", nsmap={None: SPREADSHEETML})
    worksheet.append(sheet_data)
    worksheet_xml = etree.tostring(
        worksheet, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    workbook = etree.Element(f"{_S}workbook", nsmap={None: SPREADSHEETML, "r": _R_NS})
    sheets = etree.SubElement(workbook, f"{_S}sheets")
    sheet = etree.SubElement(sheets, f"{_S}sheet")
    sheet.set("name", "Reference")
    sheet.set("sheetId", "1")
    sheet.set(f"{{{_R_NS}}}id", "rId1")
    workbook_xml = etree.tostring(workbook, xml_declaration=True, encoding="UTF-8", standalone=True)

    builder = PackageBuilder()
    builder.add_default_type("rels", _CT_RELS)
    builder.add_default_type("xml", "application/xml")
    builder.add_relationship("/", "rId1", REL_OFFICE_DOCUMENT, "xl/workbook.xml")
    builder.add_part("/xl/workbook.xml", workbook_xml, content_type=_CT_XLSX_MAIN)
    for rid, rel_type, target, part, xml, content_type in (
        (
            "rId1",
            _REL_WORKSHEET,
            "worksheets/sheet1.xml",
            "/xl/worksheets/sheet1.xml",
            worksheet_xml,
            _CT_XLSX_WORKSHEET,
        ),
        (
            "rId2",
            REL_STYLES,
            "styles.xml",
            "/xl/styles.xml",
            _XLSX_STYLES_XML,
            _CT_XLSX_STYLES,
        ),
        (
            "rId3",
            REL_THEME,
            "theme/theme1.xml",
            "/xl/theme/theme1.xml",
            _scaffold_theme_xml(),
            _CT_THEME,
        ),
        (
            "rId4",
            REL_SHARED_STRINGS,
            "sharedStrings.xml",
            "/xl/sharedStrings.xml",
            strings.to_xml(),
            _CT_XLSX_SHARED_STRINGS,
        ),
    ):
        builder.add_relationship("/xl/workbook.xml", rid, rel_type, target)
        builder.add_part(part, xml, content_type=content_type)
    builder.write(document_path)
    _normalize_zip(document_path)


def _add_xlsx_row(
    sheet_data: etree._Element,
    strings: _SharedStrings,
    number: int,
    values: list[str],
) -> None:
    row = etree.SubElement(sheet_data, f"{_S}row")
    row.set("r", str(number))
    for column, value in zip("ABCDEFGH", values, strict=False):
        cell = etree.SubElement(row, f"{_S}c")
        cell.set("r", f"{column}{number}")
        cell.set("t", "s")
        v = etree.SubElement(cell, f"{_S}v")
        v.text = str(strings.intern(value))
