"""Stable, feature-level snapshots for DOCX round-trip comparison.

Office editors legitimately rewrite relationship IDs, ZIP metadata, XML
prefixes, and application properties.  Those byte-level changes are useful
diagnostics, but they are not a document-preservation verdict.  This module
extracts the document invariants that an editor round-trip is expected to
preserve and compares them independently from the OPC package diff.
"""

from __future__ import annotations

import hashlib
import json
import posixpath
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, cast

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

NS = {"w": W, "r": R, "a": A, "pr": PKG_REL, "cp": CP, "dc": DC, "dcterms": DCTERMS}
WQ = f"{{{W}}}"
RQ = f"{{{R}}}"

FEATURES = (
    "content_blocks",
    "heading_hierarchy",
    "list_semantics",
    "table_semantics",
    "section_semantics",
    "header_footer_semantics",
    "field_codes",
    "style_semantics",
    "theme_semantics",
    "metadata_semantics",
    "security_surface",
)


@dataclass(frozen=True)
class DocxSemanticSnapshot:
    """Normalized semantic evidence extracted from one DOCX package."""

    source_sha256: str
    features: dict[str, Any]
    schema_version: int = 1

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class FeatureComparison:
    """Preservation result for one independently inspectable feature family."""

    feature: str
    preserved: bool
    base_sha256: str
    head_sha256: str


@dataclass(frozen=True)
class DocxSemanticComparison:
    """Feature matrix for a source and one editor-produced DOCX."""

    preserved: bool
    features: tuple[FeatureComparison, ...]
    changed_features: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "preserved": self.preserved,
            "changed_features": list(self.changed_features),
            "features": [asdict(feature) for feature in self.features],
        }


def _parser() -> etree.XMLParser:
    return etree.XMLParser(resolve_entities=False, no_network=True, recover=False)


def _read_xml(archive: zipfile.ZipFile, name: str) -> etree._Element | None:
    try:
        payload = archive.read(name)
    except KeyError:
        return None
    return etree.fromstring(payload, parser=_parser())


def _attr(element: etree._Element | None, name: str) -> str | None:
    if element is None:
        return None
    return cast(str | None, element.get(f"{WQ}{name}"))


def _child_val(element: etree._Element | None, child_name: str) -> str | None:
    if element is None:
        return None
    return _attr(element.find(f"{WQ}{child_name}"), "val")


def _on_off(element: etree._Element | None, child_name: str) -> bool:
    if element is None:
        return False
    child = element.find(f"{WQ}{child_name}")
    if child is None:
        return False
    return (_attr(child, "val") or "1").lower() not in {"0", "false", "off"}


def _visible_text(element: etree._Element) -> str:
    chunks: list[str] = []

    def visit(node: etree._Element, deleted: bool = False) -> None:
        is_deleted = deleted or node.tag == f"{WQ}del"
        if not is_deleted:
            if node.tag == f"{WQ}t" and node.text:
                chunks.append(node.text)
            elif node.tag == f"{WQ}tab":
                chunks.append("\t")
            elif node.tag in {f"{WQ}br", f"{WQ}cr"}:
                chunks.append("\n")
            elif node.tag == f"{WQ}noBreakHyphen":
                chunks.append("\u2011")
            elif node.tag == f"{WQ}softHyphen":
                chunks.append("\u00ad")
        for child in node:
            visit(child, is_deleted)

    visit(element)
    return "".join(chunks)


def _attributes(element: etree._Element | None, names: tuple[str, ...]) -> dict[str, str]:
    if element is None:
        return {}
    return {
        name: value
        for name in names
        if (value := element.get(f"{WQ}{name}")) is not None
    }


def _paragraph_properties(p_pr: etree._Element | None) -> dict[str, Any]:
    if p_pr is None:
        return {}
    properties: dict[str, Any] = {}
    for child in ("jc", "outlineLvl", "textDirection", "contextualSpacing"):
        value = _child_val(p_pr, child)
        if value is not None:
            properties[child] = value
    for child in ("keepNext", "keepLines", "pageBreakBefore", "widowControl"):
        if p_pr.find(f"{WQ}{child}") is not None:
            properties[child] = _on_off(p_pr, child)
    for child, attributes in {
        "spacing": ("before", "after", "line", "lineRule", "beforeAutospacing", "afterAutospacing"),
        "ind": ("left", "right", "firstLine", "hanging", "start", "end"),
    }.items():
        values = _attributes(p_pr.find(f"{WQ}{child}"), attributes)
        if values:
            properties[child] = values
    tabs = []
    for tab in p_pr.findall(f"{WQ}tabs/{WQ}tab"):
        tabs.append(_attributes(tab, ("val", "pos", "leader")))
    if tabs:
        properties["tabs"] = tabs
    return properties


def _run_properties(r_pr: etree._Element | None) -> dict[str, Any]:
    if r_pr is None:
        return {}
    properties: dict[str, Any] = {}
    fonts = _attributes(
        r_pr.find(f"{WQ}rFonts"),
        (
            "ascii",
            "hAnsi",
            "eastAsia",
            "cs",
            "asciiTheme",
            "hAnsiTheme",
            "eastAsiaTheme",
            "csTheme",
        ),
    )
    if fonts:
        properties["fonts"] = fonts
    for child in ("sz", "szCs", "u", "vertAlign"):
        value = _child_val(r_pr, child)
        if value is not None:
            properties[child] = value
    color = _attributes(r_pr.find(f"{WQ}color"), ("val", "themeColor", "themeTint", "themeShade"))
    if color:
        properties["color"] = color
    language = _attributes(r_pr.find(f"{WQ}lang"), ("val", "eastAsia", "bidi"))
    if language:
        properties["lang"] = language
    for child in ("b", "i", "strike", "caps", "smallCaps", "vanish"):
        if r_pr.find(f"{WQ}{child}") is not None:
            properties[child] = _on_off(r_pr, child)
    return properties


def _style_catalog(
    root: etree._Element | None,
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    by_id: dict[str, dict[str, Any]] = {}
    if root is None:
        return by_id, by_id
    for style in root.findall(f"{WQ}style"):
        style_id = _attr(style, "styleId")
        if not style_id:
            continue
        entry = {
            "type": _attr(style, "type"),
            "name": _child_val(style, "name"),
            "based_on": _child_val(style, "basedOn"),
            "next": _child_val(style, "next"),
            "link": _child_val(style, "link"),
            "q_format": style.find(f"{WQ}qFormat") is not None,
            "p": _paragraph_properties(style.find(f"{WQ}pPr")),
            "r": _run_properties(style.find(f"{WQ}rPr")),
        }
        by_id[style_id] = entry
    semantic = {
        style_id: entry
        for style_id, entry in sorted(by_id.items())
        if entry["type"] in {"paragraph", "character", "table"}
    }
    return by_id, semantic


def _style_outline(style_id: str | None, styles: dict[str, dict[str, Any]]) -> int | None:
    seen: set[str] = set()
    current = style_id
    while current and current not in seen:
        seen.add(current)
        entry = styles.get(current)
        if entry is None:
            break
        raw = entry.get("p", {}).get("outlineLvl")
        if raw is not None:
            try:
                return int(raw)
            except ValueError:
                return None
        current = entry.get("based_on")
    if style_id:
        compact = style_id.lower().replace(" ", "")
        if compact.startswith("heading") and compact[7:].isdigit():
            return max(0, int(compact[7:]) - 1)
    return None


def _numbering_catalog(
    root: etree._Element | None,
) -> tuple[dict[str, str], dict[str, list[dict[str, Any]]]]:
    nums: dict[str, str] = {}
    abstracts: dict[str, list[dict[str, Any]]] = {}
    if root is None:
        return nums, abstracts
    for abstract in root.findall(f"{WQ}abstractNum"):
        abstract_id = _attr(abstract, "abstractNumId")
        if abstract_id is None:
            continue
        levels = []
        for level in abstract.findall(f"{WQ}lvl"):
            levels.append(
                {
                    "level": _attr(level, "ilvl"),
                    "start": _child_val(level, "start"),
                    "format": _child_val(level, "numFmt"),
                    "text": _child_val(level, "lvlText"),
                    "suffix": _child_val(level, "suff"),
                    "alignment": _child_val(level, "lvlJc"),
                    "p": _paragraph_properties(level.find(f"{WQ}pPr")),
                    "r": _run_properties(level.find(f"{WQ}rPr")),
                }
            )
        abstracts[abstract_id] = levels
    for num in root.findall(f"{WQ}num"):
        num_id = _attr(num, "numId")
        abstract_id = _child_val(num, "abstractNumId")
        if num_id is not None and abstract_id is not None:
            nums[num_id] = abstract_id
    return nums, abstracts


def _paragraph_snapshot(
    paragraph: etree._Element,
    styles: dict[str, dict[str, Any]],
    nums: dict[str, str],
    abstracts: dict[str, list[dict[str, Any]]],
) -> dict[str, Any]:
    p_pr = paragraph.find(f"{WQ}pPr")
    style_id = _child_val(p_pr, "pStyle")
    outline_raw = _child_val(p_pr, "outlineLvl")
    outline: int | None = None
    if outline_raw is not None:
        try:
            outline = int(outline_raw)
        except ValueError:
            outline = None
    if outline is None:
        outline = _style_outline(style_id, styles)
    num_pr = p_pr.find(f"{WQ}numPr") if p_pr is not None else None
    num_id = _child_val(num_pr, "numId")
    level = _child_val(num_pr, "ilvl")
    abstract_id = nums.get(num_id or "")
    level_definition: dict[str, Any] | None = None
    if abstract_id is not None and level is not None:
        try:
            level_definition = abstracts[abstract_id][int(level)]
        except (KeyError, IndexError, ValueError):
            level_definition = None
    return {
        "text": _visible_text(paragraph),
        "style_id": style_id,
        "style_name": styles.get(style_id or "", {}).get("name"),
        "outline_level": outline,
        "paragraph_properties": _paragraph_properties(p_pr),
        "numbering": (
            {"level": level, "definition": level_definition}
            if num_id is not None
            else None
        ),
    }


def _table_snapshot(
    table: etree._Element,
    styles: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    tbl_pr = table.find(f"{WQ}tblPr")
    style_id = _child_val(tbl_pr, "tblStyle")
    rows = []
    for row in table.findall(f"{WQ}tr"):
        cells = []
        for cell in row.findall(f"{WQ}tc"):
            cells.append("\n".join(_visible_text(p) for p in cell.findall(f".//{WQ}p")))
        tr_pr = row.find(f"{WQ}trPr")
        rows.append({"header": _on_off(tr_pr, "tblHeader"), "cells": cells})
    look = _attributes(
        tbl_pr.find(f"{WQ}tblLook") if tbl_pr is not None else None,
        ("val", "firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"),
    )
    return {
        "style_id": style_id,
        "style_name": styles.get(style_id or "", {}).get("name"),
        "look": look,
        "rows": rows,
    }


def _relationship_map(root: etree._Element | None, source_part: str) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    if root is None:
        return result
    source_dir = posixpath.dirname(source_part)
    for relationship in root.findall(f"{{{PKG_REL}}}Relationship"):
        rel_id = relationship.get("Id")
        target = relationship.get("Target")
        if not rel_id or not target:
            continue
        mode = relationship.get("TargetMode", "Internal")
        normalized = (
            target
            if mode == "External"
            else posixpath.normpath(posixpath.join(source_dir, target))
        )
        result[rel_id] = {
            "type": relationship.get("Type", ""),
            "target": normalized,
            "mode": mode,
        }
    return result


def _part_content_snapshot(root: etree._Element | None) -> dict[str, Any]:
    if root is None:
        return {"text": [], "fields": []}
    paragraphs = [_visible_text(p) for p in root.findall(f".//{WQ}p")]
    fields = [
        " ".join(value.split())
        for node in root.findall(f".//{WQ}instrText") + root.findall(f".//{WQ}fldSimple")
        if (value := (node.text if node.tag == f"{WQ}instrText" else _attr(node, "instr")))
    ]
    return {"text": paragraphs, "fields": fields}


def _section_snapshots(
    archive: zipfile.ZipFile,
    document: etree._Element,
    relationships: dict[str, dict[str, str]],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    header_footer: dict[str, Any] = {}
    for index, section in enumerate(document.findall(f".//{WQ}sectPr")):
        references = []
        for kind in ("header", "footer"):
            for reference in section.findall(f"{WQ}{kind}Reference"):
                rel_id = reference.get(f"{RQ}id")
                relation = relationships.get(rel_id or "")
                reference_type = _attr(reference, "type") or "default"
                entry: dict[str, Any] = {"kind": kind, "type": reference_type}
                if relation and relation["mode"] != "External":
                    part = relation["target"]
                    content = _part_content_snapshot(_read_xml(archive, part))
                    entry["content"] = content
                    header_footer[f"section_{index}:{kind}:{reference_type}"] = content
                references.append(entry)
        columns_node = section.find(f"{WQ}cols")
        columns = _attributes(columns_node, ("num", "space", "equalWidth", "sep"))
        columns.setdefault("num", "1")
        columns.setdefault("space", "720")
        columns.setdefault("equalWidth", "1")
        columns.setdefault("sep", "0")
        if columns_node is not None:
            explicit = [
                _attributes(column, ("w", "space"))
                for column in columns_node.findall(f"{WQ}col")
            ]
            if explicit:
                columns["columns"] = explicit  # type: ignore[assignment]
        sections.append(
            {
                "type": _child_val(section, "type") or "nextPage",
                "page_size": _attributes(section.find(f"{WQ}pgSz"), ("w", "h", "orient", "code")),
                "margins": _attributes(
                    section.find(f"{WQ}pgMar"),
                    ("top", "right", "bottom", "left", "header", "footer", "gutter"),
                ),
                "columns": columns,
                "title_page": section.find(f"{WQ}titlePg") is not None,
                "references": references,
            }
        )
    return sections, header_footer


def _field_codes(document: etree._Element, header_footer: dict[str, Any]) -> list[str]:
    values: list[str] = []
    for simple in document.findall(f".//{WQ}fldSimple"):
        if value := _attr(simple, "instr"):
            values.append(" ".join(value.split()))
    for complex_field in document.findall(f".//{WQ}instrText"):
        if complex_field.text:
            values.append(" ".join(complex_field.text.split()))
    for part in header_footer.values():
        values.extend(part.get("fields", []))
    return values


def _theme_snapshot(root: etree._Element | None) -> dict[str, Any]:
    if root is None:
        return {}
    colors: dict[str, str] = {}
    scheme = root.find(".//a:clrScheme", namespaces=NS)
    if scheme is not None:
        for slot in scheme:
            if len(slot) == 0:
                continue
            color = slot[0]
            value = color.get("val") or color.get("lastClr") or ""
            colors[etree.QName(slot).localname] = value
    fonts: dict[str, str] = {}
    for family in ("majorFont", "minorFont"):
        node = root.find(f".//a:{family}", namespaces=NS)
        if node is None:
            continue
        for script in ("latin", "ea", "cs"):
            font = node.find(f"a:{script}", namespaces=NS)
            if font is not None:
                fonts[f"{family}.{script}"] = font.get("typeface", "")
    return {"colors": colors, "fonts": fonts}


def _metadata_snapshot(archive: zipfile.ZipFile) -> dict[str, Any]:
    core = _read_xml(archive, "docProps/core.xml")
    metadata: dict[str, Any] = {}
    if core is not None:
        selectors = {
            "title": "dc:title",
            "subject": "dc:subject",
            "language": "dc:language",
            "description": "dc:description",
            "keywords": "cp:keywords",
            "category": "cp:category",
            "content_status": "cp:contentStatus",
        }
        for key, selector in selectors.items():
            node = core.find(selector, namespaces=NS)
            if node is not None and node.text is not None:
                metadata[key] = node.text
    custom = _read_xml(archive, "docProps/custom.xml")
    custom_properties: dict[str, str] = {}
    if custom is not None:
        for prop in custom.findall(f"{{{CUSTOM}}}property"):
            name = prop.get("name")
            if name and len(prop):
                custom_properties[name] = "".join(prop[0].itertext())
    metadata["custom"] = custom_properties
    return metadata


def _security_surface(archive: zipfile.ZipFile) -> dict[str, Any]:
    names = set(archive.namelist())
    external = []
    for name in sorted(part for part in names if part.endswith(".rels")):
        root = _read_xml(archive, name)
        if root is None:
            continue
        for relation in root.findall(f"{{{PKG_REL}}}Relationship"):
            if relation.get("TargetMode") == "External":
                external.append(
                    {
                        "part": name,
                        "type": relation.get("Type", ""),
                        "target": relation.get("Target", ""),
                    }
                )
    return {
        "macros": sorted(name for name in names if name.lower().endswith("vbaproject.bin")),
        "embedded_objects": sorted(
            name
            for name in names
            if name.startswith(("word/embeddings/", "word/activeX/"))
        ),
        "external_relationships": external,
    }


def snapshot_docx(path: Path | str) -> DocxSemanticSnapshot:
    """Extract normalized document invariants from ``path``."""

    source = Path(path).resolve()
    payload = source.read_bytes()
    with zipfile.ZipFile(source) as archive:
        document = _read_xml(archive, "word/document.xml")
        if document is None:
            raise ValueError(f"{source} has no word/document.xml")
        styles, style_semantics = _style_catalog(_read_xml(archive, "word/styles.xml"))
        nums, abstracts = _numbering_catalog(_read_xml(archive, "word/numbering.xml"))
        relationships = _relationship_map(
            _read_xml(archive, "word/_rels/document.xml.rels"),
            "word/document.xml",
        )
        body = document.find(f"{WQ}body")
        if body is None:
            raise ValueError(f"{source} has no w:body")

        blocks: list[dict[str, Any]] = []
        tables: list[dict[str, Any]] = []
        headings: list[dict[str, Any]] = []
        lists: list[dict[str, Any]] = []
        for child in body:
            if child.tag == f"{WQ}p":
                paragraph = _paragraph_snapshot(child, styles, nums, abstracts)
                blocks.append({"kind": "paragraph", "text": paragraph["text"]})
                if paragraph["outline_level"] is not None:
                    headings.append(
                        {
                            "text": paragraph["text"],
                            "level": paragraph["outline_level"],
                            "style_name": paragraph["style_name"],
                        }
                    )
                if paragraph["numbering"] is not None:
                    lists.append(
                        {
                            "text": paragraph["text"],
                            "style_name": paragraph["style_name"],
                            **paragraph["numbering"],
                        }
                    )
            elif child.tag == f"{WQ}tbl":
                table = _table_snapshot(child, styles)
                tables.append(table)
                blocks.append(
                    {
                        "kind": "table",
                        "rows": [row["cells"] for row in table["rows"]],
                    }
                )

        sections, header_footer = _section_snapshots(archive, document, relationships)
        theme_relation = next(
            (
                relation
                for relation in relationships.values()
                if relation["type"].endswith("/theme") and relation["mode"] != "External"
            ),
            None,
        )
        theme = _theme_snapshot(
            _read_xml(archive, theme_relation["target"])
            if theme_relation is not None
            else None
        )

        features = {
            "content_blocks": blocks,
            "heading_hierarchy": headings,
            "list_semantics": lists,
            "table_semantics": tables,
            "section_semantics": sections,
            "header_footer_semantics": header_footer,
            "field_codes": _field_codes(document, header_footer),
            "style_semantics": style_semantics,
            "theme_semantics": theme,
            "metadata_semantics": _metadata_snapshot(archive),
            "security_surface": _security_surface(archive),
        }
    return DocxSemanticSnapshot(
        source_sha256=hashlib.sha256(payload).hexdigest(),
        features=features,
    )


def _feature_hash(value: Any) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def compare_docx_semantics(
    base: DocxSemanticSnapshot,
    head: DocxSemanticSnapshot,
) -> DocxSemanticComparison:
    """Compare two snapshots feature-by-feature, independent of package bytes."""

    comparisons = tuple(
        FeatureComparison(
            feature=feature,
            preserved=base.features.get(feature) == head.features.get(feature),
            base_sha256=_feature_hash(base.features.get(feature)),
            head_sha256=_feature_hash(head.features.get(feature)),
        )
        for feature in FEATURES
    )
    changed = tuple(item.feature for item in comparisons if not item.preserved)
    return DocxSemanticComparison(
        preserved=not changed,
        features=comparisons,
        changed_features=changed,
    )


__all__ = [
    "DocxSemanticComparison",
    "DocxSemanticSnapshot",
    "FEATURES",
    "FeatureComparison",
    "compare_docx_semantics",
    "snapshot_docx",
]
