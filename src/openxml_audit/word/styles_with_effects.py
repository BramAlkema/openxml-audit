"""Styles with effects validation (MS-OE376).

Validates the stylesWithEffects.xml part introduced in Office 2010.
This part mirrors styles.xml but includes DrawingML effect properties.
It uses the relationship type
  http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects
and content type
  application/vnd.ms-word.stylesWithEffects+xml
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.namespaces import REL_STYLES, REL_STYLES_WITH_EFFECTS, WORDPROCESSINGML
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
            return list(context.errors)

        # Root element must be w:styles
        expected_tag = f"{{{WORDPROCESSINGML}}}styles"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"stylesWithEffects root should be 'styles', got '{xml.tag}'",
            )
            return list(context.errors)

        self._validate_styles(xml, context)
        self._validate_consistency_with_styles(xml, package, context)
        return list(context.errors)

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
        """Check that stylesWithEffects is consistent with styles.xml.

        Word cross-checks these two parts and shows the "unreadable content"
        repair dialog when they diverge. The .NET Open XML SDK does not detect
        this; tools that modify only styles.xml (e.g. python-docx) routinely
        produce drift. We check three things:

        - styles missing from stylesWithEffects (present in styles.xml only)
        - styles missing from styles.xml (present in stylesWithEffects only)
        - w:docDefaults differing between the two parts

        We deliberately do not compare matching style definitions: by design,
        stylesWithEffects mirrors styles.xml *plus* DrawingML effects, so
        per-style content drift is expected for any style that carries an
        effect (text outline, shadow, glow, etc.) and would produce a flood
        of false positives on legitimate Word-generated files.
        """
        main_uri = package.get_main_document_uri()
        if not main_uri:
            return
        main_part = OpenXmlPart(package, main_uri)
        rel = main_part.relationships.get_first_by_type(REL_STYLES)
        if rel is None:
            return
        styles_uri = rel.resolve_target(main_uri)

        styles_xml = package.get_part_xml(styles_uri)
        if styles_xml is None:
            return

        styles_by_id = self._collect_styles_by_id(styles_xml)
        effects_by_id = self._collect_styles_by_id(effects_xml)

        # Styles in effects but not in styles.xml
        only_in_effects = sorted(set(effects_by_id) - set(styles_by_id))
        if only_in_effects:
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    f"stylesWithEffects contains {len(only_in_effects)} style(s) "
                    f"not in styles.xml: {', '.join(only_in_effects[:5])}"
                ),
                node="styleId",
                severity=ValidationSeverity.ERROR,
            )

        # Styles in styles.xml but not in effects (the python-docx failure mode)
        only_in_styles = sorted(set(styles_by_id) - set(effects_by_id))
        if only_in_styles:
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    f"styles.xml contains {len(only_in_styles)} style(s) not in "
                    f"stylesWithEffects: {', '.join(only_in_styles[:5])}"
                ),
                node="styleId",
                severity=ValidationSeverity.ERROR,
            )

        self._compare_doc_defaults(styles_xml, effects_xml, context)

    def _collect_styles_by_id(
        self, root: etree._Element
    ) -> dict[str, etree._Element]:
        result: dict[str, etree._Element] = {}
        for style in root.findall("w:style", self._ns):
            sid = style.get(f"{{{WORDPROCESSINGML}}}styleId", "").strip()
            if sid and sid not in result:
                result[sid] = style
        return result

    def _compare_doc_defaults(
        self,
        styles_xml: etree._Element,
        effects_xml: etree._Element,
        context: ValidationContext,
    ) -> None:
        styles_dd = styles_xml.find("w:docDefaults", self._ns)
        effects_dd = effects_xml.find("w:docDefaults", self._ns)

        if styles_dd is None and effects_dd is None:
            return

        if styles_dd is None or effects_dd is None:
            present = "styles.xml" if styles_dd is not None else "stylesWithEffects"
            missing = "stylesWithEffects" if styles_dd is not None else "styles.xml"
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    f"docDefaults present in {present} but missing from {missing}"
                ),
                node="docDefaults",
                severity=ValidationSeverity.ERROR,
            )
            return

        if _canonicalize(styles_dd) != _canonicalize(effects_dd):
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    "docDefaults differ between styles.xml and stylesWithEffects "
                    "(Word may treat this as unreadable content)"
                ),
                node="docDefaults",
                severity=ValidationSeverity.ERROR,
            )

def _canonicalize(elem: etree._Element) -> bytes:
    """Return a canonical byte form of `elem` for content-equality checks.

    Re-roots the subtree so namespace declarations are local before c14n2,
    which otherwise rejects elements whose namespaces are declared on an
    ancestor outside the slice being serialized.
    """
    rerooted = etree.fromstring(etree.tostring(elem))
    return cast(bytes, etree.tostring(rerooted, method="c14n2"))
