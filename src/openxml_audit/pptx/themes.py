"""PPTX theme validation.

Validates theme structure including color schemes, font schemes,
and format schemes.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import FileFormat, ValidationError
from openxml_audit.namespaces import DRAWINGML

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import ThemePart


class ThemeValidator:
    """Validates PPTX theme structure."""

    def __init__(self) -> None:
        self._ns = {"a": DRAWINGML}

    def validate(
        self, part: "ThemePart", context: "ValidationContext"
    ) -> list[ValidationError]:
        """Validate a theme part.

        Args:
            part: The theme part to validate.
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

            # Validate theme elements
            self._validate_theme_elements(xml, context)

            # Validate color scheme
            self._validate_clr_scheme(xml, context)

            # Validate font scheme
            self._validate_font_scheme(xml, context)

            # Validate format scheme
            self._validate_fmt_scheme(xml, context)

        errors.extend(context.errors)
        return errors

    def _validate_root(self, xml: etree._Element, context: "ValidationContext") -> None:
        """Validate the root theme element."""
        expected_tag = f"{{{DRAWINGML}}}theme"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Root element should be 'theme', got '{xml.tag}'",
            )

        # Theme must have a name
        name = xml.get("name")
        if not name:
            context.add_schema_error(
                "Theme missing required 'name' attribute",
                node="name",
            )

    def _validate_theme_elements(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate theme elements structure."""
        themeElements = xml.find("a:themeElements", self._ns)
        if themeElements is None:
            context.add_schema_error(
                "Theme missing required themeElements element",
            )
            return

        # Check required children of themeElements
        required = ["clrScheme", "fontScheme", "fmtScheme"]
        for elem_name in required:
            if themeElements.find(f"a:{elem_name}", self._ns) is None:
                context.add_schema_error(
                    f"themeElements missing required {elem_name} element",
                )

    def _validate_clr_scheme(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate color scheme."""
        themeElements = xml.find("a:themeElements", self._ns)
        if themeElements is None:
            return

        clrScheme = themeElements.find("a:clrScheme", self._ns)
        if clrScheme is None:
            return  # Already reported

        # Color scheme must have a name
        name = clrScheme.get("name")
        if not name:
            context.add_schema_error(
                "clrScheme missing required 'name' attribute",
                node="name",
            )

        # Required color elements
        required_colors = [
            "dk1",   # Dark 1
            "lt1",   # Light 1
            "dk2",   # Dark 2
            "lt2",   # Light 2
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hlink",  # Hyperlink
            "folHlink",  # Followed hyperlink
        ]

        for color_name in required_colors:
            color_elem = clrScheme.find(f"a:{color_name}", self._ns)
            if color_elem is None:
                context.add_schema_error(
                    f"clrScheme missing required {color_name} color",
                )
            else:
                # Each color must have a color definition child
                self._validate_color_element(color_elem, color_name, context)

    def _validate_color_element(
        self,
        color_elem: etree._Element,
        color_name: str,
        context: "ValidationContext",
    ) -> None:
        """Validate a color element has a proper color definition."""
        # Color can be defined by various elements
        valid_color_types = [
            "srgbClr",  # RGB color
            "schemeClr",  # Scheme color reference
            "sysClr",  # System color
            "prstClr",  # Preset color
            "hslClr",  # HSL color
            "scrgbClr",  # RGB percentage color
        ]

        has_color = False
        for color_type in valid_color_types:
            if color_elem.find(f"a:{color_type}", self._ns) is not None:
                has_color = True
                break

        if not has_color:
            context.add_schema_error(
                f"{color_name} element must contain a color definition",
            )

    def _validate_font_scheme(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate font scheme."""
        themeElements = xml.find("a:themeElements", self._ns)
        if themeElements is None:
            return

        fontScheme = themeElements.find("a:fontScheme", self._ns)
        if fontScheme is None:
            return  # Already reported

        # Font scheme must have a name
        name = fontScheme.get("name")
        if not name:
            context.add_schema_error(
                "fontScheme missing required 'name' attribute",
                node="name",
            )

        # Required font collections
        majorFont = fontScheme.find("a:majorFont", self._ns)
        if majorFont is None:
            context.add_schema_error(
                "fontScheme missing required majorFont element",
            )
        else:
            self._validate_font_collection(majorFont, "majorFont", context)

        minorFont = fontScheme.find("a:minorFont", self._ns)
        if minorFont is None:
            context.add_schema_error(
                "fontScheme missing required minorFont element",
            )
        else:
            self._validate_font_collection(minorFont, "minorFont", context)

    def _validate_font_collection(
        self,
        font_coll: etree._Element,
        name: str,
        context: "ValidationContext",
    ) -> None:
        """Validate a font collection (majorFont or minorFont)."""
        # Must have latin font
        latin = font_coll.find("a:latin", self._ns)
        if latin is None:
            context.add_schema_error(
                f"{name} missing required latin font",
            )
        else:
            typeface = latin.get("typeface")
            if not typeface:
                context.add_schema_error(
                    f"{name}/latin missing required 'typeface' attribute",
                    node="typeface",
                )

        # Must have east asian font
        ea = font_coll.find("a:ea", self._ns)
        if ea is None:
            context.add_schema_error(
                f"{name} missing required ea (East Asian) font",
            )

        # Must have complex script font
        cs = font_coll.find("a:cs", self._ns)
        if cs is None:
            context.add_schema_error(
                f"{name} missing required cs (Complex Script) font",
            )

    def _validate_fmt_scheme(
        self, xml: etree._Element, context: "ValidationContext"
    ) -> None:
        """Validate format scheme."""
        themeElements = xml.find("a:themeElements", self._ns)
        if themeElements is None:
            return

        fmtScheme = themeElements.find("a:fmtScheme", self._ns)
        if fmtScheme is None:
            return  # Already reported

        # Format scheme name is optional in newer Office formats
        name = fmtScheme.get("name")
        if not name and context.file_format == FileFormat.OFFICE_2007:
            context.add_schema_error(
                "fmtScheme missing required 'name' attribute",
                node="name",
            )

        # Required format lists
        required_lists = [
            ("fillStyleLst", 3),   # Fill styles - minimum 3
            ("lnStyleLst", 3),     # Line styles - minimum 3
            ("effectStyleLst", 3), # Effect styles - minimum 3
            ("bgFillStyleLst", 3), # Background fill styles - minimum 3
        ]

        for list_name, min_count in required_lists:
            style_list = fmtScheme.find(f"a:{list_name}", self._ns)
            if style_list is None:
                context.add_schema_error(
                    f"fmtScheme missing required {list_name} element",
                )
            else:
                # Count children (each represents a style)
                children = [c for c in style_list if isinstance(c.tag, str)]
                if len(children) < min_count:
                    context.add_schema_error(
                        f"{list_name} must have at least {min_count} styles, found {len(children)}",
                    )


def validate_theme(
    part: "ThemePart", context: "ValidationContext"
) -> list[ValidationError]:
    """Validate a theme part.

    Args:
        part: The theme part.
        context: Validation context.

    Returns:
        List of validation errors.
    """
    validator = ThemeValidator()
    return validator.validate(part, context)
