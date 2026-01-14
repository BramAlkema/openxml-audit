"""PPTX presentation.xml validation.

Validates the main presentation structure including slide lists,
masters, and presentation properties.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.namespaces import PRESENTATIONML, REL_SLIDE, REL_SLIDE_MASTER

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import PresentationPart


class PresentationValidator:
    """Validates PPTX presentation structure."""

    def __init__(self) -> None:
        self._ns = {"p": PRESENTATIONML}
        self._rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def validate(
        self, part: "PresentationPart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a presentation part.

        Args:
            part: The presentation part to validate.
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

            # Validate slide master list
            self._validate_slide_master_list(xml, part, context)

            # Validate slide list
            self._validate_slide_list(xml, part, context)

            # Validate notes master
            self._validate_notes_master(xml, part, context)

            # Validate handout master
            self._validate_handout_master(xml, part, context)

            # Validate slide size
            self._validate_slide_size(xml, context)

            # Validate notes size
            self._validate_notes_size(xml, context)

        errors.extend(context.errors)
        return errors

    def _validate_root(self, xml: etree._Element, context: "ValidationContext") -> None:
        """Validate the root presentation element."""
        expected_tag = f"{{{PRESENTATIONML}}}presentation"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'presentation', got '{xml.tag}'",
            )

        # autoCompress is not a valid attribute in the ECMA-376 schema
        if xml.get("autoCompress") is not None:
            context.add_schema_error(
                "Attribute 'autoCompress' is not declared.",
                node="autoCompress",
            )

    def _validate_slide_master_list(
        self,
        xml: etree._Element,
        part: "PresentationPart",
        context: "ValidationContext",
    ) -> None:
        """Validate slide master references."""
        master_list = xml.find("p:sldMasterIdLst", self._ns)

        if master_list is None:
            context.add_schema_error(
                "Presentation must have at least one slide master",
            )
            return

        master_ids = master_list.findall("p:sldMasterId", self._ns)
        if not master_ids:
            context.add_schema_error(
                "sldMasterIdLst is empty - at least one slide master required",
            )
            return

        seen_ids: set[str] = set()

        for master_id in master_ids:
            # Check for duplicate master IDs
            id_val = master_id.get("id", "")
            if id_val:
                if id_val in seen_ids:
                    context.add_semantic_error(
                        f"Duplicate slide master ID: {id_val}",
                        node="id",
                    )
                seen_ids.add(id_val)

            # Check relationship exists
            rel_id = master_id.get(f"{{{self._rel_ns}}}id", "")
            if not rel_id:
                context.add_schema_error(
                    "sldMasterId missing r:id attribute",
                    node="sldMasterId",
                )
                continue

            rel = part.relationships.get_by_id(rel_id)
            if rel is None:
                context.add_semantic_error(
                    f"Slide master relationship '{rel_id}' not found",
                    node="r:id",
                )
            elif rel.type != REL_SLIDE_MASTER:
                context.add_semantic_error(
                    f"Relationship '{rel_id}' should be slideMaster type, got '{rel.type}'",
                    node="r:id",
                )

    def _validate_slide_list(
        self,
        xml: etree._Element,
        part: "PresentationPart",
        context: "ValidationContext",
    ) -> None:
        """Validate slide references."""
        slide_list = xml.find("p:sldIdLst", self._ns)

        if slide_list is None:
            # Empty presentation - valid but possibly warn
            return

        slide_ids = slide_list.findall("p:sldId", self._ns)
        seen_ids: set[str] = set()

        for slide_id in slide_ids:
            # Check for duplicate slide IDs
            id_val = slide_id.get("id", "")
            if id_val:
                if id_val in seen_ids:
                    context.add_semantic_error(
                        f"Duplicate slide ID: {id_val}",
                        node="id",
                    )
                seen_ids.add(id_val)

            # Check relationship exists
            rel_id = slide_id.get(f"{{{self._rel_ns}}}id", "")
            if not rel_id:
                context.add_schema_error(
                    "sldId missing r:id attribute",
                    node="sldId",
                )
                continue

            rel = part.relationships.get_by_id(rel_id)
            if rel is None:
                context.add_semantic_error(
                    f"Slide relationship '{rel_id}' not found",
                    node="r:id",
                )
            elif rel.type != REL_SLIDE:
                context.add_semantic_error(
                    f"Relationship '{rel_id}' should be slide type, got '{rel.type}'",
                    node="r:id",
                )

    def _validate_notes_master(
        self,
        xml: etree._Element,
        part: "PresentationPart",
        context: "ValidationContext",
    ) -> None:
        """Validate notes master reference."""
        notes_list = xml.find("p:notesMasterIdLst", self._ns)
        if notes_list is None:
            return  # Notes master is optional

        notes_ids = notes_list.findall("p:notesMasterId", self._ns)
        if len(notes_ids) > 1:
            context.add_schema_error(
                "Only one notes master allowed",
            )

    def _validate_handout_master(
        self,
        xml: etree._Element,
        part: "PresentationPart",
        context: "ValidationContext",
    ) -> None:
        """Validate handout master reference."""
        handout_list = xml.find("p:handoutMasterIdLst", self._ns)
        if handout_list is None:
            return  # Handout master is optional

        handout_ids = handout_list.findall("p:handoutMasterId", self._ns)
        if len(handout_ids) > 1:
            context.add_schema_error(
                "Only one handout master allowed",
            )

    def _validate_slide_size(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate slide size element."""
        slide_size = xml.find("p:sldSz", self._ns)
        if slide_size is None:
            return  # Uses default size

        # Validate dimensions
        cx = slide_size.get("cx")
        cy = slide_size.get("cy")

        if cx is not None:
            try:
                cx_val = int(cx)
                if cx_val <= 0:
                    context.add_semantic_error(
                        f"Slide width must be positive, got {cx_val}",
                        node="cx",
                    )
            except ValueError:
                context.add_schema_error(
                    f"Invalid slide width: {cx}",
                    node="cx",
                )

        if cy is not None:
            try:
                cy_val = int(cy)
                if cy_val <= 0:
                    context.add_semantic_error(
                        f"Slide height must be positive, got {cy_val}",
                        node="cy",
                    )
            except ValueError:
                context.add_schema_error(
                    f"Invalid slide height: {cy}",
                    node="cy",
                )

    def _validate_notes_size(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate notes size element."""
        notes_size = xml.find("p:notesSz", self._ns)
        if notes_size is None:
            context.add_schema_error(
                "notesSz element is required",
            )
            return

        # Validate dimensions
        cx = notes_size.get("cx")
        cy = notes_size.get("cy")

        if cx is None or cy is None:
            context.add_schema_error(
                "notesSz requires cx and cy attributes",
                node="notesSz",
            )


def validate_presentation(
    part: "PresentationPart", context: "ValidationContext"
) -> list[ValidationError]:
    """Validate a presentation part.

    Args:
        part: The presentation part.
        context: Validation context.

    Returns:
        List of validation errors.
    """
    validator = PresentationValidator()
    return validator.validate(part, context)
