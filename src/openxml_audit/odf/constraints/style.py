"""Style-related ODF constraints."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType
from openxml_audit.odf._helpers import NUMBER_NS, OFFICE_NS, STYLE_NS, SVG_NS, TEXT_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule

# ── shared helpers ──────────────────────────────────────────────────────


def collect_all_style_names(*roots: etree._Element | None) -> set[str]:
    """Collect all style:name values from automatic-styles and styles containers."""
    names: set[str] = set()
    for root in roots:
        if root is None:
            continue
        for container_local in ("automatic-styles", "styles"):
            container = root.find(f"{{{OFFICE_NS}}}{container_local}")
            if container is None:
                continue
            for style in container:
                if not isinstance(style.tag, str):
                    continue
                name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                if name:
                    names.add(name)
    return names


def collect_list_style_names(*roots: etree._Element | None) -> set[str]:
    """Collect text:list-style names from automatic-styles and styles."""
    names: set[str] = set()
    for root in roots:
        if root is None:
            continue
        for container_local in ("automatic-styles", "styles"):
            container = root.find(f"{{{OFFICE_NS}}}{container_local}")
            if container is None:
                continue
            for child in container:
                if not isinstance(child.tag, str):
                    continue
                qname = etree.QName(child)
                if qname.namespace == TEXT_NS and qname.localname == "list-style":
                    name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                    if name:
                        names.add(name)
    return names


def collect_data_style_names(*roots: etree._Element | None) -> set[str]:
    """Collect number:* data style names."""
    names: set[str] = set()
    for root in roots:
        if root is None:
            continue
        for container_local in ("automatic-styles", "styles"):
            container = root.find(f"{{{OFFICE_NS}}}{container_local}")
            if container is None:
                continue
            for child in container:
                if not isinstance(child.tag, str):
                    continue
                qname = etree.QName(child)
                if qname.namespace == NUMBER_NS:
                    name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                    if name:
                        names.add(name)
    return names


def collect_page_layout_names(styles: etree._Element) -> set[str]:
    """Collect style:page-layout names from styles.xml."""
    names: set[str] = set()
    alm = styles.find(f"{{{OFFICE_NS}}}automatic-styles")
    if alm is None:
        return names
    for child in alm:
        if not isinstance(child.tag, str):
            continue
        qname = etree.QName(child)
        if qname.namespace == STYLE_NS and qname.localname == "page-layout":
            name = child.get(f"{{{STYLE_NS}}}name", "").strip()
            if name:
                names.add(name)
    return names


def collect_font_face_names(root: etree._Element) -> set[str]:
    """Collect font face declaration names from a root element."""
    names: set[str] = set()
    decls = root.find(f"{{{OFFICE_NS}}}font-face-decls")
    if decls is None:
        return names
    for child in decls:
        if not isinstance(child.tag, str):
            continue
        name = child.get(f"{{{STYLE_NS}}}name", "").strip()
        if name:
            names.add(name)
    return names


# ── constraints ─────────────────────────────────────────────────────────


class FontFaceDeclarationConstraint(OdfConstraint):
    """ODFSEMSTYLE001: font-face-decl must have svg:font-family."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSTYLE001",
            family="style",
            description="Font face declarations must have svg:font-family attribute.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part_name in ("content.xml", "styles.xml"):
            root = ctx.parsed_parts.get(part_name)
            if root is None:
                continue
            decls = root.find(f"{{{OFFICE_NS}}}font-face-decls")
            if decls is None:
                continue
            for child in decls:
                if not isinstance(child.tag, str):
                    continue
                qname = etree.QName(child)
                if qname.namespace != STYLE_NS or qname.localname != "font-face":
                    continue
                font_family = child.get(f"{{{SVG_NS}}}font-family", "").strip()
                style_name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                if not font_family:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE001",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Font face declaration '{style_name or '(unnamed)'}' in "
                                f"{part_name} is missing required svg:font-family"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors


class StyleParentRefConstraint(OdfConstraint):
    """ODFSEMSTYLE002: style:parent-style-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSTYLE002",
            family="style",
            description="Parent style references must resolve to defined styles.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        all_names = collect_all_style_names(content, styles)

        for part_name in ("content.xml", "styles.xml"):
            root = ctx.parsed_parts.get(part_name)
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    parent = style.get(f"{{{STYLE_NS}}}parent-style-name", "").strip()
                    if not parent or parent in all_names:
                        continue
                    style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE002",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Style '{style_name or '(unnamed)'}' in {part_name} "
                                f"references parent style '{parent}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors


class DataStyleRefConstraint(OdfConstraint):
    """ODFSEMSTYLE003: style:data-style-name must resolve to a number:* style."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSTYLE003",
            family="style",
            description="Data style references (style:data-style-name) must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        data_names = collect_data_style_names(content, styles)

        for part_name in ("content.xml", "styles.xml"):
            root = ctx.parsed_parts.get(part_name)
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    ref = style.get(f"{{{STYLE_NS}}}data-style-name", "").strip()
                    if not ref or ref in data_names:
                        continue
                    style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE003",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Style '{style_name or '(unnamed)'}' in {part_name} "
                                f"references data style '{ref}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors


class ListStyleRefConstraint(OdfConstraint):
    """ODFSEMSTYLE004: style:list-style-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSTYLE004",
            family="style",
            description="List style references must resolve to defined list styles.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        list_names = collect_list_style_names(content, styles)
        if not list_names and content is None and styles is None:
            return errors

        for part_name in ("content.xml", "styles.xml"):
            root = ctx.parsed_parts.get(part_name)
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    ref = style.get(f"{{{STYLE_NS}}}list-style-name", "").strip()
                    if not ref or ref in list_names:
                        continue
                    style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE004",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Style '{style_name or '(unnamed)'}' in {part_name} "
                                f"references list style '{ref}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors


class MasterPageLayoutConstraint(OdfConstraint):
    """ODFSEMSTYLE005: master-page page-layout-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSTYLE005",
            family="style",
            description="Page layout references in master pages must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        styles = ctx.parsed_parts.get("styles.xml")
        if styles is None:
            return errors

        layout_names = collect_page_layout_names(styles)
        master_pages = styles.find(f"{{{OFFICE_NS}}}master-styles")
        if master_pages is None:
            return errors

        for mp in master_pages:
            if not isinstance(mp.tag, str):
                continue
            qname = etree.QName(mp)
            if qname.namespace != STYLE_NS or qname.localname != "master-page":
                continue
            layout_ref = mp.get(f"{{{STYLE_NS}}}page-layout-name", "").strip()
            if not layout_ref or layout_ref in layout_names:
                continue
            mp_name = mp.get(f"{{{STYLE_NS}}}name", "").strip()
            errors.append(
                self._error(
                    rule_id="ODFSEMSTYLE005",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Master page '{mp_name or '(unnamed)'}' references "
                        f"page layout '{layout_ref}' which is not defined"
                    ),
                    part_uri="/styles.xml",
                )
            )
        return errors
