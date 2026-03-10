"""Styles with effects validation (MS-OE376).

Validates the stylesWithEffects.xml part introduced in Office 2010.
This part mirrors styles.xml but includes DrawingML effect properties.
It uses the relationship type
  http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects
and content type
  application/vnd.ms-word.stylesWithEffects+xml
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.namespaces import REL_STYLES_WITH_EFFECTS, WORDPROCESSINGML
from openxml_audit.parts import OpenXmlPart

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.package import OpenXmlPackage

# Expected content type
CT_STYLES_WITH_EFFECTS = "application/vnd.ms-word.stylesWithEffects+xml"

# Style types per ECMA-376 §17.7.4.17
_VALID_STYLE_TYPES = frozenset({
    "paragraph", "character", "table", "numbering",
})


class StylesWithEffectsValidator:
    """Validates the stylesWithEffects.xml part in Word documents."""

    def __init__(self) -> None:
        self._ns = {"w": WORDPROCESSINGML}

    def find_part(self, package: OpenXmlPackage) -> str | None:
        """Find the stylesWithEffects part URI via document relationships."""
        main_uri = package.get_main_document_uri()
        if not main_uri:
            return None
        main_part = OpenXmlPart(package, main_uri)
        for rel in main_part.relationships:
            if rel.type == REL_STYLES_WITH_EFFECTS:
                return rel.resolve_target(main_uri)
        return None

    def validate(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> list[ValidationError]:
        """Validate stylesWithEffects.xml if present."""
        uri = self.find_part(package)
        if uri is None:
            return []

        errors: list[ValidationError] = []
        context.set_part(OpenXmlPart(package, uri))

        # Verify content type
        ct = package.content_types.get_content_type(uri)
        if ct and ct != CT_STYLES_WITH_EFFECTS:
            context.add_schema_error(
                f"stylesWithEffects content type should be "
                f"'{CT_STYLES_WITH_EFFECTS}', got '{ct}'",
                node="ContentType",
            )

        xml = package.get_part_xml(uri)
        if xml is None:
            context.add_schema_error(
                f"stylesWithEffects part '{uri}' could not be parsed",
            )
            errors.extend(context.errors)
            return errors

        # Root element must be w:styles
        expected_tag = f"{{{WORDPROCESSINGML}}}styles"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"stylesWithEffects root should be 'styles', got '{xml.tag}'",
            )
            errors.extend(context.errors)
            return errors

        self._validate_styles(xml, context)
        self._validate_consistency_with_styles(xml, package, context)

        errors.extend(context.errors)
        return errors

    def _validate_styles(
        self, xml: etree._Element, context: ValidationContext
    ) -> None:
        """Validate individual style entries in stylesWithEffects."""
        seen_ids: set[str] = set()

        for style in xml.findall("w:style", self._ns):
            style_id = style.get(f"{{{WORDPROCESSINGML}}}styleId", "").strip()
            style_type = style.get(f"{{{WORDPROCESSINGML}}}type", "").strip()

            # styleId must not be empty
            if not style_id:
                context.add_schema_error(
                    "stylesWithEffects: style element missing styleId",
                    node="styleId",
                )
                continue

            # styleId must be unique
            if style_id in seen_ids:
                context.add_semantic_error(
                    f"stylesWithEffects: duplicate styleId '{style_id}'",
                    node="styleId",
                )
            else:
                seen_ids.add(style_id)

            # type must be valid
            if style_type and style_type not in _VALID_STYLE_TYPES:
                context.add_schema_error(
                    f"stylesWithEffects: style '{style_id}' has invalid "
                    f"type '{style_type}'",
                    node="type",
                )

            # basedOn must reference an existing style
            based_on = style.find("w:basedOn", self._ns)
            if based_on is not None:
                ref = based_on.get(f"{{{WORDPROCESSINGML}}}val", "").strip()
                if not ref:
                    context.add_schema_error(
                        f"stylesWithEffects: style '{style_id}' has empty "
                        "basedOn reference",
                        node="basedOn",
                    )

    def _validate_consistency_with_styles(
        self,
        effects_xml: etree._Element,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> None:
        """Check that stylesWithEffects styleIds are consistent with styles.xml."""
        styles_uri = None
        main_uri = package.get_main_document_uri()
        if main_uri:
            main_part = OpenXmlPart(package, main_uri)
            rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
            for rel in main_part.relationships:
                if rel.type == rel_type:
                    styles_uri = rel.resolve_target(main_uri)
                    break

        if styles_uri is None or not package.has_part(styles_uri):
            return

        styles_xml = package.get_part_xml(styles_uri)
        if styles_xml is None:
            return

        # Collect style IDs from styles.xml
        styles_ids: set[str] = set()
        for style in styles_xml.findall(
            f"{{{WORDPROCESSINGML}}}style"
        ):
            sid = style.get(f"{{{WORDPROCESSINGML}}}styleId", "").strip()
            if sid:
                styles_ids.add(sid)

        # Collect style IDs from stylesWithEffects
        effects_ids: set[str] = set()
        for style in effects_xml.findall("w:style", self._ns):
            sid = style.get(f"{{{WORDPROCESSINGML}}}styleId", "").strip()
            if sid:
                effects_ids.add(sid)

        # Styles in effects but not in styles.xml
        orphaned = effects_ids - styles_ids
        if orphaned:
            sample = sorted(orphaned)[:5]
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    f"stylesWithEffects contains {len(orphaned)} style(s) not "
                    f"in styles.xml: {', '.join(sample)}"
                ),
                node="styleId",
                severity=ValidationSeverity.WARNING,
            )
