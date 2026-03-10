"""Supertheme (theme variant) validation.

Validates the MS-OE376 themeVariantManager structure used in .thmx
supertheme packages. Based on the thememl/2012/main namespace
(MS-ODRAWXML §2.1.1910).

Structure:
  _rels/.rels → themeVariantManager rel
  themeVariants/themeVariantManager.xml → t:themeVariantManager
    t:themeVariantLst
      t:themeVariant name=... vid=... cx=... cy=... r:id=...
  themeVariants/_rels/themeVariantManager.xml.rels → variant targets
  themeVariants/variant{N}/theme/presentation.xml → variant presentation
  themeVariants/variant{N}/theme/theme/theme1.xml → variant theme
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import ValidationError
from openxml_audit.namespaces import (
    MS_THEMEML_2012,
    OFFICE_DOC_RELATIONSHIPS,
    REL_THEME_VARIANT_MANAGER,
)
from openxml_audit.parts import OpenXmlPart, ThemePart

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.package import OpenXmlPackage
    from openxml_audit.pptx.themes import ThemeValidator
    from openxml_audit.relationships import RelationshipCollection

# GUID pattern: {HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}
_GUID_RE = re.compile(
    r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}"
    r"-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
)


class SuperthemeValidator:
    """Validates supertheme (themeVariantManager) structure.

    Checks:
    - themeVariantManager.xml root element and namespace
    - themeVariantLst contains at least one themeVariant
    - Each themeVariant has required attributes (name, vid, cx, cy, r:id)
    - vid is a valid GUID
    - cx/cy are positive integers (EMU)
    - r:id references resolve to relationships
    - Variant target parts exist in package
    - Variant theme parts are valid themes
    """

    def __init__(self, theme_validator: ThemeValidator | None = None) -> None:
        self._theme_validator = theme_validator

    def find_variant_manager(
        self, package: OpenXmlPackage
    ) -> str | None:
        """Find the themeVariantManager part URI via package relationships."""
        for rel in package.relationships:
            if rel.type == REL_THEME_VARIANT_MANAGER:
                return rel.resolve_target("/")
        return None

    def validate(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> list[ValidationError]:
        """Validate supertheme structure if present."""
        manager_uri = self.find_variant_manager(package)
        if manager_uri is None:
            return []

        errors: list[ValidationError] = []
        context.set_part(OpenXmlPart(package, manager_uri))

        # Parse themeVariantManager.xml
        xml = package.get_part_xml(manager_uri)
        if xml is None:
            context.add_schema_error(
                f"themeVariantManager part '{manager_uri}' could not be parsed",
            )
            errors.extend(context.errors)
            return errors

        # Validate root element
        expected_tag = f"{{{MS_THEMEML_2012}}}themeVariantManager"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Expected root element 'themeVariantManager' in thememl/2012 "
                f"namespace, got '{xml.tag}'",
            )
            errors.extend(context.errors)
            return errors

        # Validate themeVariantLst
        ns = {"t": MS_THEMEML_2012, "r": OFFICE_DOC_RELATIONSHIPS}
        variant_lst = xml.find("t:themeVariantLst", ns)
        if variant_lst is None:
            context.add_schema_error(
                "themeVariantManager missing required themeVariantLst element",
            )
            errors.extend(context.errors)
            return errors

        # Collect variants
        variants = variant_lst.findall("t:themeVariant", ns)
        if not variants:
            context.add_schema_error(
                "themeVariantLst must contain at least one themeVariant",
            )
            errors.extend(context.errors)
            return errors

        # Load variant manager relationships
        manager_part = OpenXmlPart(package, manager_uri)
        manager_rels = manager_part.relationships

        # Validate each variant
        for i, variant in enumerate(variants):
            self._validate_variant(
                variant, i, manager_uri, manager_rels, package, context, ns
            )

        errors.extend(context.errors)
        return errors

    def _validate_variant(
        self,
        variant: etree._Element,
        index: int,
        manager_uri: str,
        manager_rels: RelationshipCollection,
        package: OpenXmlPackage,
        context: ValidationContext,
        ns: dict[str, str],
    ) -> None:
        """Validate a single themeVariant element."""
        label = f"themeVariant[{index}]"

        # Required attributes
        name = variant.get("name")
        if not name:
            context.add_schema_error(
                f"{label} missing required 'name' attribute",
                node="name",
            )

        vid = variant.get("vid")
        if not vid:
            context.add_schema_error(
                f"{label} missing required 'vid' attribute",
                node="vid",
            )
        elif not _GUID_RE.match(vid):
            context.add_schema_error(
                f"{label} 'vid' is not a valid GUID: '{vid}'",
                node="vid",
            )

        # EMU dimensions
        for attr in ("cx", "cy"):
            val = variant.get(attr)
            if val is None:
                context.add_schema_error(
                    f"{label} missing required '{attr}' attribute",
                    node=attr,
                )
            else:
                try:
                    emu = int(val)
                    if emu <= 0:
                        context.add_schema_error(
                            f"{label} '{attr}' must be positive, got {emu}",
                            node=attr,
                        )
                except ValueError:
                    context.add_schema_error(
                        f"{label} '{attr}' must be an integer, got '{val}'",
                        node=attr,
                    )

        # Relationship reference
        r_id = variant.get(f"{{{OFFICE_DOC_RELATIONSHIPS}}}id")
        if not r_id:
            context.add_semantic_error(
                f"{label} missing required r:id attribute",
                node="r:id",
            )
            return

        # Resolve relationship
        rel = manager_rels.get_by_id(r_id)
        if rel is None:
            context.add_semantic_error(
                f"{label} r:id='{r_id}' does not reference a valid relationship",
                node="r:id",
            )
            return

        # Check target part exists
        target = rel.resolve_target(manager_uri)
        if not target or not package.has_part(target):
            context.add_semantic_error(
                f"{label} relationship target '{target}' not found in package",
                node="r:id",
            )
            return

        # Validate variant theme if ThemeValidator available
        if self._theme_validator is not None and target:
            self._validate_variant_theme(
                target, package, context, label
            )

    def _validate_variant_theme(
        self,
        variant_presentation_uri: str,
        package: OpenXmlPackage,
        context: ValidationContext,
        label: str,
    ) -> None:
        """Validate the theme within a variant sub-package."""
        assert self._theme_validator is not None
        # Variant structure: variant{N}/theme/presentation.xml has rels
        # pointing to theme/theme1.xml
        variant_part = OpenXmlPart(package, variant_presentation_uri)
        theme_rels = [
            r for r in variant_part.relationships
            if r.type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        ]
        if not theme_rels:
            context.add_semantic_error(
                f"{label} variant presentation has no theme relationship",
                node="Relationship",
            )
            return

        for theme_rel in theme_rels:
            theme_target = theme_rel.resolve_target(variant_presentation_uri)
            if not theme_target or not package.has_part(theme_target):
                context.add_semantic_error(
                    f"{label} variant theme '{theme_target}' not found in package",
                    node="Relationship",
                )
                continue

            theme_part = ThemePart(package, theme_target)
            self._theme_validator.validate(theme_part, context)
