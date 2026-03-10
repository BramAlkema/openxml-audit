"""Extended and custom properties validation (docProps/app.xml, core.xml, custom.xml).

Validates OPC package metadata parts per ECMA-376 Part 2 (OPC) and
MS-OE376 extended properties.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.namespaces import (
    CORE_PROPERTIES,
    CUSTOM_PROPERTIES,
    DC,
    DCTERMS,
    EXTENDED_PROPERTIES,
    OFFICE_DOC,
    REL_CUSTOM_PROPERTIES,
    REL_EXTENDED_PROPERTIES,
    RELATIONSHIPS_METADATA_CORE,
)
from openxml_audit.parts import OpenXmlPart

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.package import OpenXmlPackage

# Extended properties (app.xml) expected element names
_APP_STRING_ELEMENTS = frozenset({
    "Application", "AppVersion", "Template", "Manager", "Company",
    "PresentationFormat", "HyperlinkBase",
})
_APP_INT_ELEMENTS = frozenset({
    "TotalTime", "Pages", "Words", "Characters", "Lines",
    "Paragraphs", "Slides", "Notes", "HiddenSlides", "MMClips",
    "CharactersWithSpaces",
})
_APP_BOOL_ELEMENTS = frozenset({
    "ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged",
})
_APP_VECTOR_ELEMENTS = frozenset({
    "HeadingPairs", "TitlesOfParts",
})
_APP_ALL_ELEMENTS = (
    _APP_STRING_ELEMENTS | _APP_INT_ELEMENTS | _APP_BOOL_ELEMENTS
    | _APP_VECTOR_ELEMENTS | {"DocSecurity", "DigSig"}
)

# Core properties (core.xml) expected element names with namespaces
_CORE_DC_ELEMENTS = frozenset({
    f"{{{DC}}}title",
    f"{{{DC}}}subject",
    f"{{{DC}}}creator",
    f"{{{DC}}}description",
    f"{{{DC}}}language",
    f"{{{DC}}}identifier",
})
_CORE_DCTERMS_ELEMENTS = frozenset({
    f"{{{DCTERMS}}}created",
    f"{{{DCTERMS}}}modified",
})
_CORE_CP_ELEMENTS = frozenset({
    f"{{{CORE_PROPERTIES}}}category",
    f"{{{CORE_PROPERTIES}}}contentStatus",
    f"{{{CORE_PROPERTIES}}}contentType",
    f"{{{CORE_PROPERTIES}}}keywords",
    f"{{{CORE_PROPERTIES}}}lastModifiedBy",
    f"{{{CORE_PROPERTIES}}}lastPrinted",
    f"{{{CORE_PROPERTIES}}}revision",
    f"{{{CORE_PROPERTIES}}}version",
})
_CORE_ALL_ELEMENTS = _CORE_DC_ELEMENTS | _CORE_DCTERMS_ELEMENTS | _CORE_CP_ELEMENTS

# VTypes namespace elements for vector validation
_VTYPE_VALID_TYPES = frozenset({
    "variant", "i1", "i2", "i4", "i8", "ui1", "ui2", "ui4", "ui8",
    "r4", "r8", "lpstr", "lpwstr", "bstr", "date", "filetime",
    "bool", "cy", "error", "blob", "oblob", "stream", "ostream",
    "storage", "ostorage", "vstream", "clsid", "cf", "null", "empty",
    "decimal", "int", "uint",
})


class PropertiesValidator:
    """Validates OPC extended/core/custom properties structure."""

    def validate(
        self,
        package: OpenXmlPackage,
        context: ValidationContext,
    ) -> list[ValidationError]:
        """Validate all property parts found in the package."""
        self._validate_core_properties(package, context)
        self._validate_extended_properties(package, context)
        self._validate_custom_properties(package, context)
        return list(context.errors)

    @staticmethod
    def _resolve_part_uri(package: OpenXmlPackage, rel_type: str) -> str | None:
        """Find a part URI via package-level relationship type."""
        rel = package.relationships.get_first_by_type(rel_type)
        return rel.resolve_target("/") if rel else None

    def _validate_core_properties(
        self, package: OpenXmlPackage, context: ValidationContext
    ) -> None:
        """Validate docProps/core.xml structure."""
        uri = self._resolve_part_uri(package, RELATIONSHIPS_METADATA_CORE)
        if uri is None:
            # Core properties are optional
            return

        context.set_part(OpenXmlPart(package, uri))
        xml = package.get_part_xml(uri)
        if xml is None:
            context.add_schema_error(
                f"Core properties part '{uri}' could not be parsed",
            )
            return

        # Root element must be cp:coreProperties
        expected_tag = f"{{{CORE_PROPERTIES}}}coreProperties"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Core properties root should be 'coreProperties', got '{xml.tag}'",
            )
            return

        # Validate children are recognized elements
        for child in xml:
            if not isinstance(child.tag, str):
                continue
            if child.tag not in _CORE_ALL_ELEMENTS:
                context.add_error(
                    error_type=ValidationErrorType.SCHEMA,
                    description=f"Unexpected element in core properties: '{child.tag}'",
                    node=etree.QName(child).localname,
                    severity=ValidationSeverity.WARNING,
                )

        # dcterms:created and dcterms:modified should have xsi:type="dcterms:W3CDTF"
        for date_tag in _CORE_DCTERMS_ELEMENTS:
            date_elem = xml.find(date_tag)
            if date_elem is not None:
                xsi_type = date_elem.get(
                    "{http://www.w3.org/2001/XMLSchema-instance}type", ""
                )
                if xsi_type and xsi_type != "dcterms:W3CDTF":
                    local = etree.QName(date_tag).localname
                    context.add_error(
                        error_type=ValidationErrorType.SCHEMA,
                        description=(
                            f"Core property '{local}' has unexpected xsi:type '{xsi_type}', "
                            "expected 'dcterms:W3CDTF'"
                        ),
                        node=local,
                        severity=ValidationSeverity.WARNING,
                    )

    def _validate_extended_properties(
        self, package: OpenXmlPackage, context: ValidationContext
    ) -> None:
        """Validate docProps/app.xml structure."""
        uri = self._resolve_part_uri(package, REL_EXTENDED_PROPERTIES)
        if uri is None:
            return

        context.set_part(OpenXmlPart(package, uri))
        xml = package.get_part_xml(uri)
        if xml is None:
            context.add_schema_error(
                f"Extended properties part '{uri}' could not be parsed",
            )
            return

        expected_tag = f"{{{EXTENDED_PROPERTIES}}}Properties"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Extended properties root should be 'Properties', got '{xml.tag}'",
            )
            return

        for child in xml:
            if not isinstance(child.tag, str):
                continue
            local = etree.QName(child).localname
            if local not in _APP_ALL_ELEMENTS:
                context.add_error(
                    error_type=ValidationErrorType.SCHEMA,
                    description=f"Unexpected element in extended properties: '{local}'",
                    node=local,
                    severity=ValidationSeverity.WARNING,
                )
                continue

            # Type validation for known elements
            if local in _APP_INT_ELEMENTS:
                self._validate_int_element(child, local, context)
            elif local in _APP_BOOL_ELEMENTS:
                self._validate_bool_element(child, local, context)
            elif local in _APP_VECTOR_ELEMENTS:
                self._validate_vector_element(child, local, context)
            elif local == "DocSecurity":
                self._validate_int_element(child, local, context)

    def _validate_int_element(
        self, elem: etree._Element, name: str, context: ValidationContext
    ) -> None:
        text = (elem.text or "").strip()
        if text:
            try:
                int(text)
            except ValueError:
                context.add_schema_error(
                    f"Extended property '{name}' should be integer, got '{text}'",
                    node=name,
                )

    def _validate_bool_element(
        self, elem: etree._Element, name: str, context: ValidationContext
    ) -> None:
        text = (elem.text or "").strip().lower()
        if text and text not in ("true", "false"):
            context.add_schema_error(
                f"Extended property '{name}' should be boolean, got '{text}'",
                node=name,
            )

    def _validate_vector_element(
        self, elem: etree._Element, name: str, context: ValidationContext
    ) -> None:
        """Validate vt:vector structure in HeadingPairs/TitlesOfParts."""
        vector = elem.find(f"{{{OFFICE_DOC}}}vector")
        if vector is None:
            context.add_schema_error(
                f"Extended property '{name}' missing vt:vector child",
                node=name,
            )
            return

        size_attr = vector.get("size", "").strip()
        if not size_attr:
            context.add_schema_error(
                f"'{name}' vt:vector missing 'size' attribute",
                node=name,
            )
        else:
            try:
                size = int(size_attr)
                actual = sum(
                    1 for c in vector if isinstance(c.tag, str)
                )
                if size != actual:
                    context.add_error(
                        error_type=ValidationErrorType.SCHEMA,
                        description=f"'{name}' vt:vector size={size} but has {actual} children",
                        node=name,
                        severity=ValidationSeverity.WARNING,
                    )
            except ValueError:
                context.add_schema_error(
                    f"'{name}' vt:vector size is not an integer: '{size_attr}'",
                    node=name,
                )

        base_type = vector.get("baseType", "").strip()
        if not base_type:
            context.add_schema_error(
                f"'{name}' vt:vector missing 'baseType' attribute",
                node=name,
            )
        elif base_type not in _VTYPE_VALID_TYPES:
            context.add_error(
                error_type=ValidationErrorType.SCHEMA,
                description=f"'{name}' vt:vector has unrecognized baseType '{base_type}'",
                node=name,
                severity=ValidationSeverity.WARNING,
            )

    def _validate_custom_properties(
        self, package: OpenXmlPackage, context: ValidationContext
    ) -> None:
        """Validate docProps/custom.xml structure."""
        uri = self._resolve_part_uri(package, REL_CUSTOM_PROPERTIES)
        if uri is None:
            return

        context.set_part(OpenXmlPart(package, uri))
        xml = package.get_part_xml(uri)
        if xml is None:
            context.add_schema_error(
                f"Custom properties part '{uri}' could not be parsed",
            )
            return

        expected_tag = f"{{{CUSTOM_PROPERTIES}}}Properties"
        if xml.tag != expected_tag:
            context.add_schema_error(
                f"Custom properties root should be 'Properties', got '{xml.tag}'",
            )
            return

        seen_names: set[str] = set()
        seen_pids: set[str] = set()

        for child in xml:
            if not isinstance(child.tag, str):
                continue
            local = etree.QName(child).localname
            if local != "property":
                context.add_schema_error(
                    f"Custom properties should only contain 'property' elements, "
                    f"found '{local}'",
                    node=local,
                )
                continue

            # fmtid is required
            fmtid = child.get("fmtid", "").strip()
            if not fmtid:
                context.add_schema_error(
                    "Custom property missing required 'fmtid' attribute",
                    node="fmtid",
                )

            # pid is required and must be unique, >= 2
            pid = child.get("pid", "").strip()
            if not pid:
                context.add_schema_error(
                    "Custom property missing required 'pid' attribute",
                    node="pid",
                )
            else:
                try:
                    pid_val = int(pid)
                    if pid_val < 2:
                        context.add_schema_error(
                            f"Custom property pid must be >= 2, got {pid_val}",
                            node="pid",
                        )
                except ValueError:
                    context.add_schema_error(
                        f"Custom property pid must be integer, got '{pid}'",
                        node="pid",
                    )
                if pid in seen_pids:
                    context.add_semantic_error(
                        f"Duplicate custom property pid '{pid}'",
                        node="pid",
                    )
                else:
                    seen_pids.add(pid)

            # name is required and should be unique
            name = child.get("name", "").strip()
            if not name:
                context.add_schema_error(
                    "Custom property missing required 'name' attribute",
                    node="name",
                )
            elif name in seen_names:
                context.add_semantic_error(
                    f"Duplicate custom property name '{name}'",
                    node="name",
                )
            else:
                seen_names.add(name)

            # Must have exactly one value child (vt:* element)
            value_children = [
                c for c in child if isinstance(c.tag, str)
            ]
            if len(value_children) == 0:
                context.add_schema_error(
                    f"Custom property '{name or '(unnamed)'}' has no value element",
                    node="property",
                )
            elif len(value_children) > 1:
                context.add_schema_error(
                    f"Custom property '{name or '(unnamed)'}' has "
                    f"{len(value_children)} value elements, expected 1",
                    node="property",
                )
