"""Cross-part reference ODF constraints."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import DRAW_NS, STYLE_NS, XLINK_NS, normalize_internal_href
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule
from openxml_audit.odf.constraints.style import collect_font_face_names


class MasterPageReferenceConstraint(OdfConstraint):
    """ODFSEMREF001/002: Presentation master-page references."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMREF001",
            family="reference",
            description="Presentation master-page references require styles.xml companion part.",
        )

    @staticmethod
    def _collect_presentation_master_refs(content: etree._Element) -> set[str]:
        refs: set[str] = set()
        pages = content.xpath(
            ".//draw:page[@draw:master-page-name]",
            namespaces={"draw": DRAW_NS},
        )
        for page in pages:
            value = page.get(f"{{{DRAW_NS}}}master-page-name", "").strip()
            if value:
                refs.add(value)
        return refs

    @staticmethod
    def _collect_master_page_definitions(styles: etree._Element) -> set[str]:
        refs: set[str] = set()
        for page in styles.xpath(
            ".//style:master-page[@style:name]",
            namespaces={"style": STYLE_NS},
        ):
            name = page.get(f"{{{STYLE_NS}}}name", "").strip()
            if name:
                refs.add(name)
        return refs

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        refs = self._collect_presentation_master_refs(content)
        if not refs:
            return errors
        if "styles.xml" not in ctx.package.manifest_paths():
            errors.append(
                self._error(
                    rule_id="ODFSEMREF001",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        "draw:master-page-name references require styles.xml "
                        "to be declared in manifest.xml"
                    ),
                    part_uri="/content.xml",
                )
            )
            return errors

        styles = ctx.parsed_parts.get("styles.xml")
        if styles is None:
            return errors
        definitions = self._collect_master_page_definitions(styles)
        for ref in sorted(refs):
            if ref in definitions:
                continue
            errors.append(
                self._error(
                    rule_id="ODFSEMREF002",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"draw:master-page-name '{ref}' does not resolve to a "
                        "style:master-page in styles.xml"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors


class FontFaceCrossPartConstraint(OdfConstraint):
    """ODFSEMREF003: font faces in content.xml should exist in styles.xml too."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMREF003",
            family="reference",
            description="Font face declarations in content.xml must match styles.xml.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        styles = ctx.parsed_parts.get("styles.xml")
        if content is None or styles is None:
            return errors
        if "styles.xml" not in ctx.package.manifest_paths():
            return errors

        content_fonts = collect_font_face_names(content)
        styles_fonts = collect_font_face_names(styles)

        for font in sorted(content_fonts - styles_fonts):
            errors.append(
                self._error(
                    rule_id="ODFSEMREF003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Font face '{font}' declared in content.xml but not in styles.xml"
                    ),
                    part_uri="/content.xml",
                    severity=ValidationSeverity.WARNING,
                )
            )
        return errors


class EmbeddedObjectRefConstraint(OdfConstraint):
    """ODFSEMREF004: embedded object xlink:href must resolve in package."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMREF004",
            family="reference",
            description="Embedded object xlink:href references must resolve in package.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        zip_members = ctx.package.zip_members()
        manifest_paths = ctx.package.manifest_paths()
        reported: set[str] = set()

        for obj in content.iter(f"{{{DRAW_NS}}}object", f"{{{DRAW_NS}}}object-ole"):
            href = normalize_internal_href(obj.get(f"{{{XLINK_NS}}}href", ""))
            if href is None:
                continue
            if href in reported:
                continue

            found = (
                href in zip_members
                or href.rstrip("/") + "/content.xml" in zip_members
                or href in manifest_paths
                or href.rstrip("/") + "/" in manifest_paths
            )
            if not found:
                reported.add(href)
                errors.append(
                    self._error(
                        rule_id="ODFSEMREF004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Embedded object href '{href}' does not resolve in package"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class ImageRefConstraint(OdfConstraint):
    """ODFSEMREF005: image xlink:href must resolve in package."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMREF005",
            family="reference",
            description="Image xlink:href references must resolve in package.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        zip_members = ctx.package.zip_members()
        reported: set[str] = set()

        for img in content.iter(f"{{{DRAW_NS}}}image"):
            href = normalize_internal_href(img.get(f"{{{XLINK_NS}}}href", ""))
            if href is None:
                continue
            if href in reported:
                continue
            if href not in zip_members:
                reported.add(href)
                errors.append(
                    self._error(
                        rule_id="ODFSEMREF005",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Image href '{href}' does not resolve in package"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors
