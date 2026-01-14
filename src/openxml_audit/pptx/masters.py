"""PPTX slide master and layout validation.

Validates slide masters, slide layouts, and their relationships.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import ValidationError
from openxml_audit.namespaces import DRAWINGML, PRESENTATIONML

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import OpenXmlPart, SlideLayoutPart, SlideMasterPart


class MasterValidator:
    """Validates PPTX slide master and layout structure."""

    def __init__(self) -> None:
        self._ns = {
            "p": PRESENTATIONML,
            "a": DRAWINGML,
        }
        self._rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def validate_master(
        self, part: "SlideMasterPart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a slide master part.

        Args:
            part: The slide master part to validate.
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
            self._validate_master_root(xml, context)

            # Validate common slide data
            self._validate_cSld(xml, "sldMaster", context)

            # Validate color map
            self._validate_clr_map(xml, context)

            # Validate slide layout list
            self._validate_sld_layout_id_lst(xml, part, context)

            # Validate theme relationship
            self._validate_theme_relationship(part, context)

            # Validate text styles
            self._validate_tx_styles(xml, context)

        errors.extend(context.errors)
        return errors

    def validate_layout(
        self, part: "SlideLayoutPart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a slide layout part.

        Args:
            part: The slide layout part to validate.
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
            self._validate_layout_root(xml, context)

            # Validate common slide data
            self._validate_cSld(xml, "sldLayout", context)

            # Validate color map override
            self._validate_clr_map_ovr(xml, context)

            # Validate slide master relationship
            self._validate_master_relationship(part, context)

        errors.extend(context.errors)
        return errors

    def _validate_master_root(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate the slide master root element."""
        expected_tag = f"{{{PRESENTATIONML}}}sldMaster"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'sldMaster', got '{xml.tag}'",
            )

    def _validate_layout_root(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate the slide layout root element."""
        expected_tag = f"{{{PRESENTATIONML}}}sldLayout"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'sldLayout', got '{xml.tag}'",
            )

        # Check layout type
        layout_type = xml.get("type")
        if layout_type:
            valid_types = {
                "title", "tx", "twoColTx", "tbl", "txAndChart",
                "chartAndTx", "dgm", "chart", "txAndClipArt",
                "clipArtAndTx", "titleOnly", "blank", "txAndObj",
                "objAndTx", "objOnly", "obj", "txAndMedia",
                "mediaAndTx", "objOverTx", "txOverObj", "txAndTwoObj",
                "twoObjAndTx", "twoObjOverTx", "fourObj", "vertTx",
                "clipArtAndVertTx", "vertTitleAndTx", "vertTitleAndTxOverChart",
                "twoObj", "objAndTwoObj", "twoObjAndObj", "cust",
                "secHead", "twoTxTwoObj", "objTx", "picTx",
            }
            if layout_type not in valid_types:
                context.add_semantic_error(
                    f"Invalid layout type: '{layout_type}'",
                    node="type",
                )

    def _validate_cSld(
        self,
        xml: etree._Element,
        parent_type: str,
        context: "ValidationContext",
    ) -> None:
        """Validate common slide data."""
        cSld = xml.find("p:cSld", self._ns)
        if cSld is None:
            context.add_schema_error(
                f"{parent_type} missing required cSld element",
            )
            return

        # Validate shape tree
        spTree = cSld.find("p:spTree", self._ns)
        if spTree is None:
            context.add_schema_error(
                "cSld missing required spTree element",
            )
            return

        # Validate shape tree structure
        nvGrpSpPr = spTree.find("p:nvGrpSpPr", self._ns)
        if nvGrpSpPr is None:
            context.add_schema_error(
                "spTree missing required nvGrpSpPr element",
            )

        grpSpPr = spTree.find("p:grpSpPr", self._ns)
        if grpSpPr is None:
            context.add_schema_error(
                "spTree missing required grpSpPr element",
            )

    def _validate_clr_map(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate color map in slide master."""
        clrMap = xml.find("p:clrMap", self._ns)
        if clrMap is None:
            context.add_schema_error(
                "sldMaster missing required clrMap element",
            )
            return

        # Required color map attributes
        required_attrs = [
            "bg1", "tx1", "bg2", "tx2",
            "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink",
        ]

        for attr in required_attrs:
            if clrMap.get(attr) is None:
                context.add_schema_error(
                    f"clrMap missing required '{attr}' attribute",
                    node=attr,
                )

    def _validate_clr_map_ovr(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate color map override in slide layout."""
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

    def _validate_sld_layout_id_lst(
        self,
        xml: etree._Element,
        part: "SlideMasterPart",
        context: "ValidationContext",
    ) -> None:
        """Validate slide layout references in master."""
        layoutIdLst = xml.find("p:sldLayoutIdLst", self._ns)
        if layoutIdLst is None:
            context.add_schema_error(
                "sldMaster missing required sldLayoutIdLst element",
            )
            return

        layoutIds = layoutIdLst.findall("p:sldLayoutId", self._ns)
        if not layoutIds:
            context.add_schema_error(
                "sldLayoutIdLst is empty - at least one layout required",
            )
            return

        seen_ids: set[str] = set()
        layout_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"

        for layoutId in layoutIds:
            # Check for duplicate IDs
            id_val = layoutId.get("id", "")
            if id_val:
                if id_val in seen_ids:
                    context.add_semantic_error(
                        f"Duplicate slide layout ID: {id_val}",
                        node="id",
                    )
                seen_ids.add(id_val)

            # Check relationship
            rel_id = layoutId.get(f"{{{self._rel_ns}}}id", "")
            if not rel_id:
                context.add_schema_error(
                    "sldLayoutId missing r:id attribute",
                    node="sldLayoutId",
                )
                continue

            rel = part.relationships.get_by_id(rel_id)
            if rel is None:
                context.add_semantic_error(
                    f"Slide layout relationship '{rel_id}' not found",
                    node="r:id",
                )
            elif rel.type != layout_rel_type:
                context.add_semantic_error(
                    f"Relationship '{rel_id}' should be slideLayout type",
                    node="r:id",
                )

    def _validate_theme_relationship(
        self, part: "SlideMasterPart", context: "ValidationContext"
    ) -> None:
        """Validate slide master has a theme relationship."""
        theme_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

        theme_rels = list(part.relationships.get_by_type(theme_type))
        if not theme_rels:
            context.add_semantic_error(
                "Slide master must have a theme relationship",
            )
        elif len(theme_rels) > 1:
            context.add_semantic_error(
                f"Slide master has {len(theme_rels)} theme relationships, expected 1",
            )

    def _validate_master_relationship(
        self, part: "SlideLayoutPart", context: "ValidationContext"
    ) -> None:
        """Validate slide layout has a master relationship."""
        master_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"

        master_rels = list(part.relationships.get_by_type(master_type))
        if not master_rels:
            context.add_semantic_error(
                "Slide layout must have a slideMaster relationship",
            )
        elif len(master_rels) > 1:
            context.add_semantic_error(
                f"Slide layout has {len(master_rels)} slideMaster relationships, expected 1",
            )

    def _validate_tx_styles(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate text styles in slide master."""
        txStyles = xml.find("p:txStyles", self._ns)
        if txStyles is None:
            return  # Optional

        # Check for expected text style elements
        expected_styles = ["titleStyle", "bodyStyle", "otherStyle"]
        for style_name in expected_styles:
            style = txStyles.find(f"p:{style_name}", self._ns)
            if style is None:
                # Not required, but if txStyles exists, typically has these
                pass


def validate_slide_master(
    part: "SlideMasterPart", context: "ValidationContext"
) -> list[ValidationError]:
    """Validate a slide master part.

    Args:
        part: The slide master part.
        context: Validation context.

    Returns:
        List of validation errors.
    """
    validator = MasterValidator()
    return validator.validate_master(part, context)


def validate_slide_layout(
    part: "SlideLayoutPart", context: "ValidationContext"
) -> list[ValidationError]:
    """Validate a slide layout part.

    Args:
        part: The slide layout part.
        context: Validation context.

    Returns:
        List of validation errors.
    """
    validator = MasterValidator()
    return validator.validate_layout(part, context)
