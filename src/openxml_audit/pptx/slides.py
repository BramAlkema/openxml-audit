"""PPTX slide validation.

Validates individual slides including shape trees, text bodies,
and embedded content.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import ValidationError
from openxml_audit.namespaces import DRAWINGML, PRESENTATIONML

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import SlidePart


class SlideValidator:
    """Validates PPTX slide structure."""

    def __init__(self) -> None:
        self._ns = {
            "p": PRESENTATIONML,
            "a": DRAWINGML,
        }
        self._rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def validate(
        self, part: "SlidePart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a slide part.

        Args:
            part: The slide part to validate.
            context: The validation context.

        Returns:
            List of validation errors.
        """
        context.set_part(part)
        errors: list[ValidationError] = []

        xml = part.xml
        if xml is None:
            return errors

        with ElementContext(context, xml):
            # Validate root element
            self._validate_root(xml, context)

            # Validate common slide data
            self._validate_cSld(xml, context)

            # Validate color map override
            self._validate_clr_map_ovr(xml, context)

            # Validate slide layout relationship
            self._validate_layout_relationship(part, context)

        errors.extend(context.errors)
        return errors

    def _validate_root(self, xml: etree._Element, context: "ValidationContext") -> None:
        """Validate the root slide element."""
        expected_tag = f"{{{PRESENTATIONML}}}sld"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'sld', got '{xml.tag}'",
            )

    def _validate_cSld(self, xml: etree._Element, context: "ValidationContext") -> None:
        """Validate common slide data."""
        cSld = xml.find("p:cSld", self._ns)
        if cSld is None:
            context.add_schema_error(
                "Slide missing required cSld (common slide data) element",
            )
            return

        # Validate shape tree
        spTree = cSld.find("p:spTree", self._ns)
        if spTree is None:
            context.add_schema_error(
                "cSld missing required spTree (shape tree) element",
            )
            return

        self._validate_shape_tree(spTree, context)

    def _validate_shape_tree(
        self, spTree: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate the shape tree structure."""
        # Validate group shape properties
        nvGrpSpPr = spTree.find("p:nvGrpSpPr", self._ns)
        if nvGrpSpPr is None:
            context.add_schema_error(
                "spTree missing required nvGrpSpPr element",
            )
        else:
            self._validate_nv_grp_sp_pr(nvGrpSpPr, context)

        grpSpPr = spTree.find("p:grpSpPr", self._ns)
        if grpSpPr is None:
            context.add_schema_error(
                "spTree missing required grpSpPr element",
            )

        # Validate shapes
        seen_ids: set[str] = set()

        for child in spTree:
            tag = child.tag
            if isinstance(tag, str):
                local_name = tag.split("}")[-1] if "}" in tag else tag

                if local_name in ("sp", "grpSp", "graphicFrame", "cxnSp", "pic"):
                    self._validate_shape(child, local_name, seen_ids, context)

    def _validate_nv_grp_sp_pr(
        self, nvGrpSpPr: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate non-visual group shape properties."""
        cNvPr = nvGrpSpPr.find("p:cNvPr", self._ns)
        if cNvPr is None:
            context.add_schema_error(
                "nvGrpSpPr missing required cNvPr element",
            )
        else:
            # Validate required attributes
            if cNvPr.get("id") is None:
                context.add_schema_error(
                    "cNvPr missing required 'id' attribute",
                    node="id",
                )
            if cNvPr.get("name") is None:
                context.add_schema_error(
                    "cNvPr missing required 'name' attribute",
                    node="name",
                )

        cNvGrpSpPr = nvGrpSpPr.find("p:cNvGrpSpPr", self._ns)
        if cNvGrpSpPr is None:
            context.add_schema_error(
                "nvGrpSpPr missing required cNvGrpSpPr element",
            )

        nvPr = nvGrpSpPr.find("p:nvPr", self._ns)
        if nvPr is None:
            context.add_schema_error(
                "nvGrpSpPr missing required nvPr element",
            )

    def _validate_shape(
        self,
        shape: etree._Element,
        shape_type: str,
        seen_ids: set[str],
        context: "ValidationContext",
    ) -> None:
        """Validate a shape element."""
        # Find non-visual properties based on shape type
        nv_prop_map = {
            "sp": "nvSpPr",
            "grpSp": "nvGrpSpPr",
            "graphicFrame": "nvGraphicFramePr",
            "cxnSp": "nvCxnSpPr",
            "pic": "nvPicPr",
        }

        nv_prop_name = nv_prop_map.get(shape_type)
        if nv_prop_name:
            nv_prop = shape.find(f"p:{nv_prop_name}", self._ns)
            if nv_prop is not None:
                cNvPr = nv_prop.find("p:cNvPr", self._ns)
                if cNvPr is not None:
                    shape_id = cNvPr.get("id")
                    if shape_id:
                        if shape_id in seen_ids:
                            context.add_semantic_error(
                                f"Duplicate shape ID: {shape_id}",
                                node="id",
                            )
                        seen_ids.add(shape_id)

        # Validate shape properties exist
        sp_prop_map = {
            "sp": "spPr",
            "grpSp": "grpSpPr",
            "graphicFrame": "xfrm",
            "cxnSp": "spPr",
            "pic": "spPr",
        }

        sp_prop_name = sp_prop_map.get(shape_type)
        if sp_prop_name and shape_type != "graphicFrame":
            sp_prop = shape.find(f"p:{sp_prop_name}", self._ns)
            if sp_prop is None:
                context.add_schema_error(
                    f"{shape_type} missing required {sp_prop_name} element",
                )

        # Validate text body for shapes that can have text
        if shape_type == "sp":
            self._validate_text_body(shape, context)

    def _validate_text_body(
        self, shape: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate shape text body."""
        txBody = shape.find("p:txBody", self._ns)
        if txBody is None:
            return  # Text body is optional

        # Check body properties
        bodyPr = txBody.find("a:bodyPr", self._ns)
        if bodyPr is None:
            context.add_schema_error(
                "txBody missing required bodyPr element",
            )

        # Check for at least one paragraph
        paragraphs = txBody.findall("a:p", self._ns)
        if not paragraphs:
            context.add_schema_error(
                "txBody must have at least one paragraph (a:p)",
            )

    def _validate_clr_map_ovr(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate color map override."""
        clrMapOvr = xml.find("p:clrMapOvr", self._ns)
        if clrMapOvr is None:
            return  # Optional

        # Should have either masterClrMapping or overrideClrMapping
        has_master = clrMapOvr.find("a:masterClrMapping", self._ns) is not None
        has_override = clrMapOvr.find("a:overrideClrMapping", self._ns) is not None

        if not has_master and not has_override:
            context.add_schema_error(
                "clrMapOvr must have masterClrMapping or overrideClrMapping",
            )

    def _validate_layout_relationship(
        self, part: "SlidePart", context: "ValidationContext"
    ) -> None:
        """Validate slide has a layout relationship."""
        layout_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"

        layout_rels = list(part.relationships.get_by_type(layout_type))
        if not layout_rels:
            context.add_semantic_error(
                "Slide must have a slideLayout relationship",
            )
        elif len(layout_rels) > 1:
            context.add_semantic_error(
                f"Slide has {len(layout_rels)} slideLayout relationships, expected 1",
            )


def validate_slide(
    part: "SlidePart", context: "ValidationContext"
) -> list[ValidationError]:
    """Validate a slide part.

    Args:
        part: The slide part.
        context: Validation context.

    Returns:
        List of validation errors.
    """
    validator = SlideValidator()
    return validator.validate(part, context)
