"""Drawing validation constraints (M6).

Rules for shape attributes, group nesting, connector resolution,
and custom shape geometry in ODF documents.
"""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import (
    DR3D_NS,
    DRAW_NS,
    SVG_NS,
    XLINK_NS,
    normalize_internal_href,
)
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule
from openxml_audit.odf.constraints.style import collect_all_style_names

# Draw shape element local names
_SHAPE_LOCAL_NAMES = frozenset({
    "rect", "circle", "ellipse", "line", "polyline", "polygon",
    "path", "custom-shape", "frame", "connector", "caption",
    "measure", "regular-polygon", "g",
})

# Position/size attributes (SVG namespace)
_POSITION_ATTRS = (f"{{{SVG_NS}}}x", f"{{{SVG_NS}}}y")
_SIZE_ATTRS = (f"{{{SVG_NS}}}width", f"{{{SVG_NS}}}height")


def _is_draw_shape(tag: str) -> bool:
    """Check whether a tag is a draw: shape element."""
    if not isinstance(tag, str):
        return False
    qname = etree.QName(tag)
    return qname.namespace == DRAW_NS and qname.localname in _SHAPE_LOCAL_NAMES


def _iter_draw_shapes(root: etree._Element) -> list[etree._Element]:
    """Collect all draw shape elements from a document tree."""
    return [el for el in root.iter() if _is_draw_shape(el.tag)]


def _get_draw_shapes(ctx: EvaluationContext) -> list[etree._Element]:
    """Cached draw shape collection from content.xml."""
    content = ctx.parsed_parts.get("content.xml")
    if content is None:
        return []
    return ctx.cached("draw_shapes", lambda: _iter_draw_shapes(content))  # type: ignore[return-value]


class DrawShapePositionConstraint(OdfConstraint):
    """ODFSEMDRAW001: Non-group shapes should have position/size attributes."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW001",
            family="drawing",
            description="Non-group draw shapes should have position and size attributes.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for shape in _get_draw_shapes(ctx):
            qname = etree.QName(shape)
            # Groups (draw:g), lines, and connectors don't require x/y/width/height
            if qname.localname in ("g", "line", "connector"):
                continue
            missing: list[str] = []
            for attr in (*_POSITION_ATTRS, *_SIZE_ATTRS):
                if shape.get(attr) is None:
                    missing.append(etree.QName(attr).localname)
            if missing:
                name = shape.get(f"{{{DRAW_NS}}}name", "(unnamed)")
                errors.append(
                    self._error(
                        rule_id="ODFSEMDRAW001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Shape '{name}' ({qname.localname}) is missing "
                            f"position/size attributes: {', '.join(missing)}"
                        ),
                        part_uri="/content.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors


class DrawGroupNestingConstraint(OdfConstraint):
    """ODFSEMDRAW002: draw:g groups must not be nested beyond a reasonable depth."""

    MAX_DEPTH = 32

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW002",
            family="drawing",
            description="Shape group nesting must not exceed maximum depth.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        group_tag = f"{{{DRAW_NS}}}g"

        def check_depth(elem: etree._Element, depth: int) -> None:
            if depth > self.MAX_DEPTH:
                name = elem.get(f"{{{DRAW_NS}}}name", "(unnamed)")
                errors.append(
                    self._error(
                        rule_id="ODFSEMDRAW002",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Shape group '{name}' nesting depth {depth} "
                            f"exceeds maximum ({self.MAX_DEPTH})"
                        ),
                        part_uri="/content.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
                return
            for child in elem:
                if isinstance(child.tag, str) and child.tag == group_tag:
                    check_depth(child, depth + 1)

        for elem in content.iter(group_tag):
            # Only check top-level groups (those whose parent is not a group)
            parent = elem.getparent()
            if parent is not None and isinstance(parent.tag, str) and parent.tag == group_tag:
                continue
            check_depth(elem, 1)

        return errors


class DrawConnectorResolveConstraint(OdfConstraint):
    """ODFSEMDRAW003: Connector start/end shape references must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW003",
            family="drawing",
            description="Connector start-shape and end-shape must reference existing shapes.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        # Collect all shape IDs (xml:id and draw:id)
        shape_ids: set[str] = set()
        for shape in _get_draw_shapes(ctx):
            for attr in (
                "{http://www.w3.org/XML/1998/namespace}id",
                f"{{{DRAW_NS}}}id",
            ):
                val = shape.get(attr, "").strip()
                if val:
                    shape_ids.add(val)

        # Check connectors
        reported: set[str] = set()
        for conn in content.iter(f"{{{DRAW_NS}}}connector"):
            for attr_local in ("start-shape", "end-shape"):
                ref = conn.get(f"{{{DRAW_NS}}}{attr_local}", "").strip()
                if ref and ref not in shape_ids and ref not in reported:
                    reported.add(ref)
                    errors.append(
                        self._error(
                            rule_id="ODFSEMDRAW003",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Connector draw:{attr_local}='{ref}' "
                                "does not reference an existing shape"
                            ),
                            part_uri="/content.xml",
                        )
                    )
        return errors


class DrawCustomGeometryConstraint(OdfConstraint):
    """ODFSEMDRAW004: Custom shapes must have draw:enhanced-geometry child."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW004",
            family="drawing",
            description="Custom shapes must contain draw:enhanced-geometry.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        geom_tag = f"{{{DRAW_NS}}}enhanced-geometry"
        for shape in content.iter(f"{{{DRAW_NS}}}custom-shape"):
            has_geom = any(
                isinstance(child.tag, str) and child.tag == geom_tag
                for child in shape
            )
            if not has_geom:
                name = shape.get(f"{{{DRAW_NS}}}name", "(unnamed)")
                errors.append(
                    self._error(
                        rule_id="ODFSEMDRAW004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Custom shape '{name}' is missing "
                            "draw:enhanced-geometry child element"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class DrawFrameHrefConstraint(OdfConstraint):
    """ODFSEMDRAW005: draw:frame with xlink:href must reference existing parts."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW005",
            family="drawing",
            description="Frame xlink:href references must resolve to package parts.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        members = ctx.package.zip_members()
        reported: set[str] = set()

        for frame in content.iter(f"{{{DRAW_NS}}}frame"):
            for child in frame:
                if not isinstance(child.tag, str):
                    continue
                raw_href = child.get(f"{{{XLINK_NS}}}href", "")
                href = normalize_internal_href(raw_href)
                if href is None:
                    continue
                if href not in members and href not in reported:
                    reported.add(href)
                    errors.append(
                        self._error(
                            rule_id="ODFSEMDRAW005",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Frame child references '{href}' "
                                "which is not present in the package"
                            ),
                            part_uri="/content.xml",
                        )
                    )
        return errors


class Draw3dSceneConstraint(OdfConstraint):
    """ODFSEMDRAW006: dr3d:scene must contain at least one 3D shape."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW006",
            family="drawing",
            description="3D scenes must contain at least one 3D shape element.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        dr3d_shapes = frozenset({
            f"{{{DR3D_NS}}}cube",
            f"{{{DR3D_NS}}}sphere",
            f"{{{DR3D_NS}}}extrude",
            f"{{{DR3D_NS}}}rotate",
            f"{{{DR3D_NS}}}scene",
        })

        for scene in content.iter(f"{{{DR3D_NS}}}scene"):
            has_shape = any(
                isinstance(child.tag, str) and child.tag in dr3d_shapes
                for child in scene
            )
            if not has_shape:
                errors.append(
                    self._error(
                        rule_id="ODFSEMDRAW006",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="dr3d:scene contains no 3D shape elements",
                        part_uri="/content.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors


class DrawStyleRefConstraint(OdfConstraint):
    """ODFSEMDRAW007: draw:style-name references must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMDRAW007",
            family="drawing",
            description="Shape draw:style-name references must resolve to defined styles.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        styles = ctx.parsed_parts.get("styles.xml")
        all_names = collect_all_style_names(content, styles)

        reported: set[str] = set()
        for shape in _get_draw_shapes(ctx):
            ref = shape.get(f"{{{DRAW_NS}}}style-name", "").strip()
            if ref and ref not in all_names and ref not in reported:
                reported.add(ref)
                name = shape.get(f"{{{DRAW_NS}}}name", "(unnamed)")
                errors.append(
                    self._error(
                        rule_id="ODFSEMDRAW007",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Shape '{name}' references style '{ref}' "
                            "which is not defined"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors
