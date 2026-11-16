"""Presentation-related ODF constraints."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType
from openxml_audit.odf._helpers import (
    DRAW_NS,
    OFFICE_NS,
    PRESENTATION_NS,
    STYLE_NS,
    XLINK_NS,
    normalize_internal_href,
)
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule


class PresentationPageNameConstraint(OdfConstraint):
    """ODFSEMPRES001: Presentation page names must be present and unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES001",
            family="presentation",
            description="Presentation page names must be present and unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            name = page.get(f"{{{DRAW_NS}}}name", "").strip()
            if not name:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="Presentation page is missing required draw:name",
                        part_uri="/content.xml",
                    )
                )
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate presentation page name '{name}'",
                        part_uri="/content.xml",
                    )
                )
                continue
            seen.add(name)
        return errors


class PresentationMinPagesConstraint(OdfConstraint):
    """ODFSEMPRES002: Presentation must contain at least one draw:page."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES002",
            family="presentation",
            description="Presentation must contain at least one draw:page.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return []

        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            return []
        pres = body.find(f"{{{OFFICE_NS}}}presentation")
        if pres is None:
            return []

        pages = pres.findall(f"{{{DRAW_NS}}}page")
        if pages:
            return []

        return [
            self._error(
                rule_id="ODFSEMPRES002",
                error_type=ValidationErrorType.SEMANTIC,
                description="Presentation body contains no draw:page elements",
                part_uri="/content.xml",
            )
        ]


class PresentationPageLayoutConstraint(OdfConstraint):
    """ODFSEMPRES003: Presentation page layout references must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES003",
            family="presentation",
            description="Presentation page layout references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        if content is None:
            return errors

        layout_names: set[str] = set()
        for source in (content, styles):
            if source is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = source.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for child in container:
                    if not isinstance(child.tag, str):
                        continue
                    qname = etree.QName(child)
                    if (
                        qname.namespace == STYLE_NS
                        and qname.localname == "presentation-page-layout"
                    ):
                        name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                        if name:
                            layout_names.add(name)

        if not layout_names:
            return errors

        reported: set[str] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            ref = page.get(f"{{{PRESENTATION_NS}}}presentation-page-layout-name", "").strip()
            if not ref or ref in layout_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMPRES003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=f"Presentation page layout '{ref}' is not defined",
                    part_uri="/content.xml",
                )
            )
        return errors


# ── New M2 rules ────────────────────────────────────────────────────────


class CustomShowSlideRefConstraint(OdfConstraint):
    """ODFSEMPRES004: Custom show slide references must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES004",
            family="presentation",
            description="Custom show slide references must resolve to existing pages.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        page_names: set[str] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            name = page.get(f"{{{DRAW_NS}}}name", "").strip()
            if name:
                page_names.add(name)

        for show in content.iter(f"{{{PRESENTATION_NS}}}show"):
            pages_attr = show.get(f"{{{PRESENTATION_NS}}}pages", "").strip()
            show_name = show.get(f"{{{PRESENTATION_NS}}}name", "").strip()
            if not pages_attr:
                continue
            for ref in pages_attr.split(","):
                ref = ref.strip()
                if ref and ref not in page_names:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMPRES004",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Custom show '{show_name or '(unnamed)'}' references "
                                f"page '{ref}' which does not exist"
                            ),
                            part_uri="/content.xml",
                        )
                    )
        return errors


class DrawLayerUniqueConstraint(OdfConstraint):
    """ODFSEMPRES005: Draw layer names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES005",
            family="presentation",
            description="Draw layer names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for layer in content.iter(f"{{{DRAW_NS}}}layer"):
            name = layer.get(f"{{{DRAW_NS}}}name", "").strip()
            if not name:
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES005",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate draw layer name '{name}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(name)
        return errors


class SoundHrefConstraint(OdfConstraint):
    """ODFSEMPRES006: Presentation sound xlink:href must resolve in package."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES006",
            family="presentation",
            description="Sound xlink:href references must resolve in package.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        zip_members = ctx.package.zip_members()
        reported: set[str] = set()

        for sound in content.iter(f"{{{PRESENTATION_NS}}}sound"):
            href = normalize_internal_href(sound.get(f"{{{XLINK_NS}}}href", ""))
            if href is None:
                continue
            if href in reported:
                continue
            if href not in zip_members:
                reported.add(href)
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES006",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(f"Sound href '{href}' does not resolve in package"),
                        part_uri="/content.xml",
                    )
                )
        return errors


class HeaderFooterDeclUniqueConstraint(OdfConstraint):
    """ODFSEMPRES007: Header/footer/date-time declaration names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES007",
            family="presentation",
            description="Header/footer/date-time declaration names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        decl_tags = (
            f"{{{PRESENTATION_NS}}}header-decl",
            f"{{{PRESENTATION_NS}}}footer-decl",
            f"{{{PRESENTATION_NS}}}date-time-decl",
        )
        for tag in decl_tags:
            seen: set[str] = set()
            for decl in content.iter(tag):
                name = decl.get(f"{{{PRESENTATION_NS}}}name", "").strip()
                if not name:
                    continue
                if name in seen:
                    local = etree.QName(decl).localname
                    errors.append(
                        self._error(
                            rule_id="ODFSEMPRES007",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=f"Duplicate {local} name '{name}'",
                            part_uri="/content.xml",
                        )
                    )
                else:
                    seen.add(name)
        return errors


class HeaderFooterDeclRefConstraint(OdfConstraint):
    """ODFSEMPRES008: Page header/footer/date-time refs must resolve to declarations."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES008",
            family="presentation",
            description="Page header/footer/date-time references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        # Collect declared names by type
        decl_map: dict[str, set[str]] = {
            "header": set(),
            "footer": set(),
            "date-time": set(),
        }
        for decl in content.iter(f"{{{PRESENTATION_NS}}}header-decl"):
            name = decl.get(f"{{{PRESENTATION_NS}}}name", "").strip()
            if name:
                decl_map["header"].add(name)
        for decl in content.iter(f"{{{PRESENTATION_NS}}}footer-decl"):
            name = decl.get(f"{{{PRESENTATION_NS}}}name", "").strip()
            if name:
                decl_map["footer"].add(name)
        for decl in content.iter(f"{{{PRESENTATION_NS}}}date-time-decl"):
            name = decl.get(f"{{{PRESENTATION_NS}}}name", "").strip()
            if name:
                decl_map["date-time"].add(name)

        if not any(decl_map.values()):
            return errors

        ref_attrs = (
            (f"{{{PRESENTATION_NS}}}use-header-name", "header", decl_map["header"]),
            (f"{{{PRESENTATION_NS}}}use-footer-name", "footer", decl_map["footer"]),
            (
                f"{{{PRESENTATION_NS}}}use-date-time-name",
                "date-time",
                decl_map["date-time"],
            ),
        )

        reported: set[tuple[str, str]] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            for attr, kind, defined in ref_attrs:
                ref = page.get(attr, "").strip()
                if not ref or ref in defined or (kind, ref) in reported:
                    continue
                reported.add((kind, ref))
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES008",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Page references {kind} declaration '{ref}' which is not defined"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class PresentationSettingsConstraint(OdfConstraint):
    """ODFSEMPRES009: Presentation settings first-page must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES009",
            family="presentation",
            description="Presentation settings first-page must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        page_names: set[str] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            name = page.get(f"{{{DRAW_NS}}}name", "").strip()
            if name:
                page_names.add(name)

        for settings in content.iter(f"{{{PRESENTATION_NS}}}settings"):
            first = settings.get(f"{{{PRESENTATION_NS}}}start-page", "").strip()
            if first and first not in page_names:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES009",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Presentation settings start-page '{first}' "
                            "does not reference an existing page"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class AnimationTargetConstraint(OdfConstraint):
    """ODFSEMPRES010: Animation target element must exist on the page."""

    ANIM_NS = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
    SMIL_NS = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES010",
            family="presentation",
            description="Animation target elements must exist.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        # Collect all xml:id and draw:id values
        all_ids: set[str] = set()
        xml_id_attr = "{http://www.w3.org/XML/1998/namespace}id"
        for elem in content.iter():
            if not isinstance(elem.tag, str):
                continue
            for attr in (xml_id_attr, f"{{{DRAW_NS}}}id"):
                val = elem.get(attr, "").strip()
                if val:
                    all_ids.add(val)

        if not all_ids:
            return errors

        reported: set[str] = set()
        for anim in content.iter():
            if not isinstance(anim.tag, str):
                continue
            qname = etree.QName(anim)
            if qname.namespace != self.ANIM_NS:
                continue
            target = anim.get(f"{{{self.SMIL_NS}}}targetElement", "").strip()
            if not target or target in all_ids or target in reported:
                continue
            reported.add(target)
            errors.append(
                self._error(
                    rule_id="ODFSEMPRES010",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(f"Animation targets element '{target}' which does not exist"),
                    part_uri="/content.xml",
                )
            )
        return errors


class TransitionTypeConstraint(OdfConstraint):
    """ODFSEMPRES011: Slide transition type must be a valid value."""

    VALID_TYPES = frozenset(
        {
            "manual",
            "automatic",
            "semi-automatic",
        }
    )

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES011",
            family="presentation",
            description="Slide transition type must be valid.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            trans_type = page.get(f"{{{PRESENTATION_NS}}}transition-type", "").strip()
            if not trans_type:
                continue
            if trans_type not in self.VALID_TYPES:
                page_name = page.get(f"{{{DRAW_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES011",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Page '{page_name or '(unnamed)'}' has invalid "
                            f"transition-type '{trans_type}'"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class NotesPageRefConstraint(OdfConstraint):
    """ODFSEMPRES012: Presentation notes pages must correspond to existing pages."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES012",
            family="presentation",
            description="Notes pages must correspond to existing draw pages.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        # Collect draw:page names
        page_names: set[str] = set()
        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            return errors
        pres = body.find(f"{{{OFFICE_NS}}}presentation")
        if pres is None:
            return errors

        for page in pres.findall(f"{{{DRAW_NS}}}page"):
            name = page.get(f"{{{DRAW_NS}}}name", "").strip()
            if name:
                page_names.add(name)

        # Check presentation:notes draw:page references in styles.xml
        styles = ctx.parsed_parts.get("styles.xml")
        if styles is None:
            return errors

        for notes in styles.iter(f"{{{PRESENTATION_NS}}}notes"):
            page_ref = notes.get(f"{{{DRAW_NS}}}name", "").strip()
            if page_ref and page_ref not in page_names:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES012",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Notes page '{page_ref}' does not correspond to an existing draw:page"
                        ),
                        part_uri="/styles.xml",
                    )
                )
        return errors


class PresentationClassConstraint(OdfConstraint):
    """ODFSEMPRES013: presentation:class attribute must be a valid value."""

    VALID_CLASSES = frozenset(
        {
            "title",
            "outline",
            "subtitle",
            "text",
            "graphic",
            "object",
            "chart",
            "table",
            "orgchart",
            "page",
            "notes",
            "handout",
            "header",
            "footer",
            "date-time",
            "page-number",
        }
    )

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMPRES013",
            family="presentation",
            description="Presentation placeholder class must be valid.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        reported: set[str] = set()
        for elem in content.iter():
            if not isinstance(elem.tag, str):
                continue
            cls = elem.get(f"{{{PRESENTATION_NS}}}class", "").strip()
            if not cls or cls in self.VALID_CLASSES or cls in reported:
                continue
            reported.add(cls)
            errors.append(
                self._error(
                    rule_id="ODFSEMPRES013",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=f"Invalid presentation:class value '{cls}'",
                    part_uri="/content.xml",
                )
            )
        return errors
