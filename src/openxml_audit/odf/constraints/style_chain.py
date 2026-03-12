"""Style inheritance chain validation constraints (M3).

These rules detect the #1 source of ODF interoperability failures:
broken style chains, orphaned styles, cycles, and missing defaults.
"""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import (
    DRAW_NS,
    OFFICE_NS,
    STYLE_NS,
    TABLE_NS,
    TEXT_NS,
)
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule

# ── helpers ─────────────────────────────────────────────────────────────


def _iter_styles(
    *roots: etree._Element | None,
) -> list[tuple[str, etree._Element]]:
    """Yield (part_name_hint, style_element) for all style containers."""
    result: list[tuple[str, etree._Element]] = []
    for idx, root in enumerate(roots):
        if root is None:
            continue
        part = "content.xml" if idx == 0 else "styles.xml"
        for container_local in ("automatic-styles", "styles"):
            container = root.find(f"{{{OFFICE_NS}}}{container_local}")
            if container is None:
                continue
            for child in container:
                if isinstance(child.tag, str):
                    result.append((part, child))
    return result


def _build_parent_map(
    *roots: etree._Element | None,
) -> dict[str, str]:
    """Build name -> parent-style-name mapping for all styles."""
    parents: dict[str, str] = {}
    for _, style in _iter_styles(*roots):
        name = style.get(f"{{{STYLE_NS}}}name", "").strip()
        parent = style.get(f"{{{STYLE_NS}}}parent-style-name", "").strip()
        if name and parent:
            parents[name] = parent
    return parents


def _collect_style_names_by_family(
    *roots: etree._Element | None,
) -> dict[str, set[str]]:
    """Collect style names grouped by style:family."""
    by_family: dict[str, set[str]] = {}
    for _, style in _iter_styles(*roots):
        name = style.get(f"{{{STYLE_NS}}}name", "").strip()
        family = style.get(f"{{{STYLE_NS}}}family", "").strip()
        if name and family:
            by_family.setdefault(family, set()).add(name)
    return by_family


def _collect_automatic_style_names(
    root: etree._Element | None,
) -> set[str]:
    """Collect style names from automatic-styles only."""
    names: set[str] = set()
    if root is None:
        return names
    container = root.find(f"{{{OFFICE_NS}}}automatic-styles")
    if container is None:
        return names
    for child in container:
        if not isinstance(child.tag, str):
            continue
        name = child.get(f"{{{STYLE_NS}}}name", "").strip()
        if name:
            names.add(name)
    return names


def _collect_referenced_style_names(root: etree._Element) -> set[str]:
    """Collect all style names referenced in body content."""
    refs: set[str] = set()
    style_attrs = (
        f"{{{TEXT_NS}}}style-name",
        f"{{{TABLE_NS}}}style-name",
        f"{{{DRAW_NS}}}style-name",
        f"{{{TABLE_NS}}}default-cell-style-name",
    )
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        for attr in style_attrs:
            val = elem.get(attr, "").strip()
            if val:
                refs.add(val)
    return refs


# ── constraints ─────────────────────────────────────────────────────────


class StyleCycleConstraint(OdfConstraint):
    """ODFSEMCHAIN001: Style parent chains must not contain cycles."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN001",
            family="style-chain",
            description="Style parent chains must not contain cycles.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        parents = _build_parent_map(content, styles)

        reported: set[str] = set()
        for name in parents:
            visited: set[str] = set()
            current = name
            while current in parents and current not in visited:
                visited.add(current)
                current = parents[current]
            if current in visited and current not in reported:
                reported.add(current)
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Style inheritance cycle detected involving '{current}'"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class OrphanedAutoStyleConstraint(OdfConstraint):
    """ODFSEMCHAIN002: Automatic styles should be referenced in the document body."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN002",
            family="style-chain",
            description="Automatic styles should be referenced in the document body.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        auto_names = _collect_automatic_style_names(content)
        if not auto_names:
            return errors

        referenced = _collect_referenced_style_names(content)

        # Also count parent references — auto styles can reference each other
        parents = _build_parent_map(content)
        for parent in parents.values():
            referenced.add(parent)

        # Auto styles with a parent-style-name are part of a style chain
        # and should not be flagged as orphaned
        for name in list(auto_names):
            if name in parents:
                referenced.add(name)

        orphaned = sorted(auto_names - referenced)
        for name in orphaned:
            errors.append(
                self._error(
                    rule_id="ODFSEMCHAIN002",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Automatic style '{name}' is declared but never referenced"
                    ),
                    part_uri="/content.xml",
                    severity=ValidationSeverity.WARNING,
                )
            )
        return errors


class DefaultStyleFamilyConstraint(OdfConstraint):
    """ODFSEMCHAIN003: Each used style family should have a default-style."""

    COMMON_FAMILIES = ("paragraph", "text", "table", "table-cell", "table-row", "graphic")

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN003",
            family="style-chain",
            description="Used style families should have a default-style defined.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        styles = ctx.parsed_parts.get("styles.xml")
        if styles is None:
            return errors

        # Collect families that have a default-style
        default_families: set[str] = set()
        styles_container = styles.find(f"{{{OFFICE_NS}}}styles")
        if styles_container is not None:
            for child in styles_container:
                if not isinstance(child.tag, str):
                    continue
                qname = etree.QName(child)
                if qname.namespace == STYLE_NS and qname.localname == "default-style":
                    family = child.get(f"{{{STYLE_NS}}}family", "").strip()
                    if family:
                        default_families.add(family)

        # Only flag missing default-styles for families where automatic
        # styles in content.xml use that family without an explicit
        # parent-style-name.  Those styles implicitly inherit from the
        # default-style, so a missing default-style would be problematic.
        # Named styles (in office:styles) and automatic styles with an
        # explicit parent don't rely on the default-style fallback.
        content = ctx.parsed_parts.get("content.xml")
        needs_default: set[str] = set()
        if content is not None:
            auto_container = content.find(f"{{{OFFICE_NS}}}automatic-styles")
            if auto_container is not None:
                for child in auto_container:
                    if not isinstance(child.tag, str):
                        continue
                    family = child.get(f"{{{STYLE_NS}}}family", "").strip()
                    parent = child.get(f"{{{STYLE_NS}}}parent-style-name", "").strip()
                    if family and not parent:
                        needs_default.add(family)

        for family in self.COMMON_FAMILIES:
            if family in needs_default and family not in default_families:
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN003",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Style family '{family}' is used but has no "
                            "default-style in styles.xml"
                        ),
                        part_uri="/styles.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors


class StyleMapTargetConstraint(OdfConstraint):
    """ODFSEMCHAIN004: style:map condition-target styles must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN004",
            family="style-chain",
            description="Conditional style:map targets must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")

        from openxml_audit.odf.constraints.style import collect_all_style_names

        all_names = collect_all_style_names(content, styles)

        reported: set[str] = set()
        for _, style in _iter_styles(content, styles):
            for child in style:
                if not isinstance(child.tag, str):
                    continue
                qname = etree.QName(child)
                if qname.namespace != STYLE_NS or qname.localname != "map":
                    continue
                target = child.get(f"{{{STYLE_NS}}}apply-style-name", "").strip()
                if not target or target in all_names or target in reported:
                    continue
                reported.add(target)
                style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Style '{style_name or '(unnamed)'}' has style:map "
                            f"targeting '{target}' which is not defined"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class MasterPageHeaderFooterConstraint(OdfConstraint):
    """ODFSEMCHAIN005: Master-page header/footer must reference valid page-layout."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN005",
            family="style-chain",
            description="Master page with header/footer must have a valid page-layout.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        styles = ctx.parsed_parts.get("styles.xml")
        if styles is None:
            return errors

        from openxml_audit.odf.constraints.style import collect_page_layout_names

        layout_names = collect_page_layout_names(styles)
        master_styles = styles.find(f"{{{OFFICE_NS}}}master-styles")
        if master_styles is None:
            return errors

        for mp in master_styles:
            if not isinstance(mp.tag, str):
                continue
            qname = etree.QName(mp)
            if qname.namespace != STYLE_NS or qname.localname != "master-page":
                continue

            has_header_footer = False
            for child in mp:
                if not isinstance(child.tag, str):
                    continue
                cq = etree.QName(child)
                if cq.namespace == STYLE_NS and cq.localname in (
                    "header", "footer", "header-left", "footer-left",
                ):
                    has_header_footer = True
                    break

            if not has_header_footer:
                continue

            layout_ref = mp.get(f"{{{STYLE_NS}}}page-layout-name", "").strip()
            if not layout_ref:
                mp_name = mp.get(f"{{{STYLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN005",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Master page '{mp_name or '(unnamed)'}' has header/footer "
                            "but no page-layout-name"
                        ),
                        part_uri="/styles.xml",
                    )
                )
            elif layout_ref not in layout_names:
                mp_name = mp.get(f"{{{STYLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN005",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Master page '{mp_name or '(unnamed)'}' references "
                            f"page-layout '{layout_ref}' which is not defined"
                        ),
                        part_uri="/styles.xml",
                    )
                )
        return errors


class StyleFamilyMismatchConstraint(OdfConstraint):
    """ODFSEMCHAIN006: Child style family must match parent style family."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN006",
            family="style-chain",
            description="Child and parent style must have the same style:family.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")

        # Build name -> family mapping
        name_family: dict[str, str] = {}
        for _, style in _iter_styles(content, styles):
            name = style.get(f"{{{STYLE_NS}}}name", "").strip()
            family = style.get(f"{{{STYLE_NS}}}family", "").strip()
            if name and family:
                name_family[name] = family

        for _, style in _iter_styles(content, styles):
            name = style.get(f"{{{STYLE_NS}}}name", "").strip()
            parent = style.get(f"{{{STYLE_NS}}}parent-style-name", "").strip()
            if not name or not parent:
                continue
            child_family = style.get(f"{{{STYLE_NS}}}family", "").strip()
            parent_family = name_family.get(parent, "")
            if child_family and parent_family and child_family != parent_family:
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN006",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Style '{name}' (family '{child_family}') inherits from "
                            f"'{parent}' (family '{parent_family}') — family mismatch"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class StyleNameEmptyConstraint(OdfConstraint):
    """ODFSEMCHAIN007: Styles must have a non-empty style:name."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN007",
            family="style-chain",
            description="Styles must have a non-empty style:name attribute.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")

        for part, style in _iter_styles(content, styles):
            qname = etree.QName(style)
            if qname.namespace != STYLE_NS or qname.localname != "style":
                continue
            name = style.get(f"{{{STYLE_NS}}}name", "").strip()
            if not name:
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN007",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"style:style in {part} has empty or missing style:name"
                        ),
                        part_uri=self._normalize_part_uri(part),
                    )
                )
        return errors


class StyleDuplicateNameConstraint(OdfConstraint):
    """ODFSEMCHAIN008: Named styles within the same scope must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN008",
            family="style-chain",
            description="Style names must be unique within their scope.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")

        # Check within each part + container scope
        for root, part_name in ((content, "content.xml"), (styles, "styles.xml")):
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                seen: set[str] = set()
                for child in container:
                    if not isinstance(child.tag, str):
                        continue
                    name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                    if not name:
                        continue
                    if name in seen:
                        errors.append(
                            self._error(
                                rule_id="ODFSEMCHAIN008",
                                error_type=ValidationErrorType.SEMANTIC,
                                description=(
                                    f"Duplicate style name '{name}' in "
                                    f"{part_name}/{container_local}"
                                ),
                                part_uri=self._normalize_part_uri(part_name),
                            )
                        )
                    else:
                        seen.add(name)
        return errors


class DeepInheritanceConstraint(OdfConstraint):
    """ODFSEMCHAIN009: Style inheritance chains should not be excessively deep."""

    MAX_DEPTH = 20

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN009",
            family="style-chain",
            description="Style inheritance chains must not exceed maximum depth.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        parents = _build_parent_map(content, styles)

        reported: set[str] = set()
        for name in parents:
            depth = 0
            current = name
            visited: set[str] = set()
            while current in parents and current not in visited:
                visited.add(current)
                current = parents[current]
                depth += 1
            if depth > self.MAX_DEPTH and name not in reported:
                reported.add(name)
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHAIN009",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Style '{name}' has inheritance depth {depth} "
                            f"(max {self.MAX_DEPTH})"
                        ),
                        part_uri="/content.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors


class NextStyleRefConstraint(OdfConstraint):
    """ODFSEMCHAIN010: style:next-style-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN010",
            family="style-chain",
            description="Next style references must resolve to defined styles.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")

        from openxml_audit.odf.constraints.style import collect_all_style_names

        all_names = collect_all_style_names(content, styles)

        reported: set[str] = set()
        for part, style in _iter_styles(content, styles):
            ref = style.get(f"{{{STYLE_NS}}}next-style-name", "").strip()
            if not ref or ref in all_names or ref in reported:
                continue
            reported.add(ref)
            name = style.get(f"{{{STYLE_NS}}}name", "").strip()
            errors.append(
                self._error(
                    rule_id="ODFSEMCHAIN010",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Style '{name or '(unnamed)'}' references "
                        f"next-style '{ref}' which is not defined"
                    ),
                    part_uri=self._normalize_part_uri(part),
                )
            )
        return errors


class MasterPageNextRefConstraint(OdfConstraint):
    """ODFSEMCHAIN011: Master-page next-style-name must resolve to another master-page."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN011",
            family="style-chain",
            description="Master page next-style-name must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        styles = ctx.parsed_parts.get("styles.xml")
        if styles is None:
            return errors

        master_styles = styles.find(f"{{{OFFICE_NS}}}master-styles")
        if master_styles is None:
            return errors

        mp_names: set[str] = set()
        for mp in master_styles:
            if not isinstance(mp.tag, str):
                continue
            qname = etree.QName(mp)
            if qname.namespace == STYLE_NS and qname.localname == "master-page":
                name = mp.get(f"{{{STYLE_NS}}}name", "").strip()
                if name:
                    mp_names.add(name)

        for mp in master_styles:
            if not isinstance(mp.tag, str):
                continue
            qname = etree.QName(mp)
            if qname.namespace != STYLE_NS or qname.localname != "master-page":
                continue
            next_ref = mp.get(f"{{{STYLE_NS}}}next-style-name", "").strip()
            if not next_ref or next_ref in mp_names:
                continue
            mp_name = mp.get(f"{{{STYLE_NS}}}name", "").strip()
            errors.append(
                self._error(
                    rule_id="ODFSEMCHAIN011",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Master page '{mp_name or '(unnamed)'}' references "
                        f"next-style-name '{next_ref}' which is not a defined master-page"
                    ),
                    part_uri="/styles.xml",
                )
            )
        return errors


class FontFamilyConsistencyConstraint(OdfConstraint):
    """ODFSEMCHAIN012: Font-face declarations used in styles must be declared."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHAIN012",
            family="style-chain",
            description="Font names used in styles must be declared in font-face-decls.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")

        from openxml_audit.odf.constraints.style import collect_font_face_names

        declared: set[str] = set()
        for root in (content, styles):
            if root is not None:
                declared |= collect_font_face_names(root)

        if not declared:
            return errors

        # Collect font-name references in style properties
        reported: set[str] = set()
        for _, style in _iter_styles(content, styles):
            for prop in style:
                if not isinstance(prop.tag, str):
                    continue
                for attr in (
                    f"{{{STYLE_NS}}}font-name",
                    f"{{{STYLE_NS}}}font-name-asian",
                    f"{{{STYLE_NS}}}font-name-complex",
                ):
                    ref = prop.get(attr, "").strip()
                    if ref and ref not in declared and ref not in reported:
                        reported.add(ref)
                        errors.append(
                            self._error(
                                rule_id="ODFSEMCHAIN012",
                                error_type=ValidationErrorType.SEMANTIC,
                                description=(
                                    f"Font name '{ref}' used in style properties "
                                    "but not declared in font-face-decls"
                                ),
                                part_uri="/content.xml",
                                severity=ValidationSeverity.WARNING,
                            )
                        )
        return errors
