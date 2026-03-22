"""Security-oriented checks for OOXML relationships and active content."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from urllib.parse import urlsplit

from lxml import etree

from openxml_audit.context import ElementContext, ValidationContext
from openxml_audit.errors import ValidationErrorType, ValidationSeverity
from openxml_audit.namespaces import (
    CONTENT_TYPES,
    PRESENTATIONML,
    RELATIONSHIPS,
    SPREADSHEETML,
    WORDPROCESSINGML,
)
from openxml_audit.package import OpenXmlPackage
from openxml_audit.parts import OpenXmlPart
from openxml_audit.semantic.attributes import SemanticConstraint

DANGEROUS_SCHEMES = ("javascript:", "data:", "vbscript:", "file:")
SSRF_TARGETS = (
    "169.254.169.254",
    "metadata.google",
    "metadata.azure",
    "localhost",
    "127.0.0.1",
    "0.0.0.0",
    "::1",
    "[::1]",
)

ACTIVE_CONTENT_TYPES = frozenset(
    {
        "application/vnd.ms-office.activeX",
        "application/vnd.ms-office.activeX+xml",
        "application/vnd.ms-office.vbaProject",
        "application/vnd.ms-office.vbaProjectSignature",
        "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml",
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
        "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
        "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabled.main+xml",
        "application/vnd.ms-excel.addin.macroEnabled.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.oleObject",
    }
)

ACTIVE_CONTENT_TAGS = frozenset(
    {
        f"{{{PRESENTATIONML}}}control",
        f"{{{PRESENTATIONML}}}embeddedFont",
        f"{{{PRESENTATIONML}}}oleObj",
        f"{{{WORDPROCESSINGML}}}control",
        f"{{{WORDPROCESSINGML}}}object",
        f"{{{SPREADSHEETML}}}control",
        f"{{{SPREADSHEETML}}}oleObject",
    }
)

_RELATIONSHIP_TAG = f"{{{RELATIONSHIPS}}}Relationship"
_CONTENT_TYPES_PATH = "/[Content_Types].xml"
_CONTENT_TYPES_DEFAULT = f"{{{CONTENT_TYPES}}}Default"
_CONTENT_TYPES_OVERRIDE = f"{{{CONTENT_TYPES}}}Override"


def _truncate(value: str, limit: int = 80) -> str:
    text = value.strip()
    if len(text) <= limit:
        return text
    return text[: limit - 3] + "..."


def _parse_part_xml(package: OpenXmlPackage, part_uri: str) -> etree._Element | None:
    content = package.get_part_content(part_uri)
    if content is None:
        return None
    try:
        return etree.fromstring(content)
    except etree.XMLSyntaxError:
        return None


def _walk_elements(
    element: etree._Element,
    context: ValidationContext,
    visitor: Callable[[etree._Element, ValidationContext], None],
) -> None:
    if context.should_stop:
        return

    with ElementContext(context, element):
        visitor(element, context)
        for child in element:
            if isinstance(child.tag, str):
                _walk_elements(child, context, visitor)


def _looks_like_ssrf_target(target: str) -> bool:
    lowered = target.strip().lower()
    if any(marker in lowered for marker in SSRF_TARGETS):
        return True

    parsed = urlsplit(lowered)
    hostname = (parsed.hostname or "").lower()
    return hostname in {marker.strip("[]") for marker in SSRF_TARGETS}


def _local_name(element: etree._Element) -> str:
    if not isinstance(element.tag, str):
        return ""
    return etree.QName(element).localname


@dataclass(frozen=True)
class DangerousUriConstraint(SemanticConstraint):
    """Flag relationship targets with dangerous schemes or SSRF-like hosts."""

    def validate(self, element: etree._Element, context: ValidationContext) -> bool:
        target = (element.get("Target") or "").strip()
        if not target:
            return True

        target_lower = target.lower()
        if target_lower.startswith(DANGEROUS_SCHEMES):
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=f"Relationship target uses dangerous URI scheme: {_truncate(target)}",
                node="Target",
                related_node=element.get("Id"),
                severity=ValidationSeverity.ERROR,
                error_id="Sec_DangerousUri",
            )
            return False

        if _looks_like_ssrf_target(target):
            context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    "Relationship target references an internal or cloud metadata endpoint: "
                    f"{_truncate(target)}"
                ),
                node="Target",
                related_node=element.get("Id"),
                severity=ValidationSeverity.WARNING,
                error_id="Sec_SsrfUri",
            )
            return False

        return True


@dataclass(frozen=True)
class ExternalRelationshipConstraint(SemanticConstraint):
    """Warn on external relationships so callers can audit fetched content."""

    def validate(self, element: etree._Element, context: ValidationContext) -> bool:
        if (element.get("TargetMode") or "Internal") != "External":
            return True

        target = (element.get("Target") or "").strip()
        context.add_error(
            error_type=ValidationErrorType.SEMANTIC,
            description=f"External relationship target: {_truncate(target)}",
            node="TargetMode",
            related_node=element.get("Id"),
            severity=ValidationSeverity.WARNING,
            error_id="Sec_ExternalRelationship",
        )
        return False


def validate_relationship_security(
    package: OpenXmlPackage,
    rels_part_uri: str,
    context: ValidationContext,
    dangerous_uri_constraint: DangerousUriConstraint | None = None,
    external_relationship_constraint: ExternalRelationshipConstraint | None = None,
) -> None:
    """Scan a .rels part for dangerous or external targets."""
    rels_root = _parse_part_xml(package, rels_part_uri)
    if rels_root is None:
        return

    dangerous_uri = dangerous_uri_constraint or DangerousUriConstraint()
    external_relationship = external_relationship_constraint or ExternalRelationshipConstraint()

    context.set_part(OpenXmlPart(package, rels_part_uri))

    def visitor(element: etree._Element, inner_context: ValidationContext) -> None:
        if element.tag != _RELATIONSHIP_TAG:
            return
        dangerous_uri.validate(element, inner_context)
        external_relationship.validate(element, inner_context)

    _walk_elements(rels_root, context, visitor)


def validate_active_content(package: OpenXmlPackage, context: ValidationContext) -> None:
    """Detect macro-enabled content types and active-content XML elements."""
    _validate_active_content_types(package, context)
    if context.should_stop:
        return
    _validate_active_content_elements(package, context)


def _validate_active_content_types(
    package: OpenXmlPackage,
    context: ValidationContext,
) -> None:
    root = _parse_part_xml(package, _CONTENT_TYPES_PATH)
    if root is None:
        return

    context.set_part(OpenXmlPart(package, _CONTENT_TYPES_PATH))

    def visitor(element: etree._Element, inner_context: ValidationContext) -> None:
        if element.tag not in {_CONTENT_TYPES_DEFAULT, _CONTENT_TYPES_OVERRIDE}:
            return

        content_type = (element.get("ContentType") or "").strip()
        if content_type not in ACTIVE_CONTENT_TYPES:
            return

        if element.tag == _CONTENT_TYPES_OVERRIDE:
            part_name = (element.get("PartName") or "").strip()
            description = (
                f"Active content content type '{content_type}' declared for part '{part_name}'"
            )
            node = "Override"
        else:
            extension = (element.get("Extension") or "").strip()
            suffix = f".{extension}" if extension else "<unknown>"
            description = (
                f"Active content content type '{content_type}' declared for extension '{suffix}'"
            )
            node = "Default"

        inner_context.add_error(
            error_type=ValidationErrorType.SEMANTIC,
            description=description,
            node=node,
            severity=ValidationSeverity.ERROR,
            error_id="Sec_ActiveContentType",
        )

    _walk_elements(root, context, visitor)


def _validate_active_content_elements(
    package: OpenXmlPackage,
    context: ValidationContext,
) -> None:
    for part_uri in package.list_parts():
        content_type = package.content_types.get_content_type(part_uri)
        if not content_type or "xml" not in content_type:
            continue

        root = _parse_part_xml(package, part_uri)
        if root is None:
            continue

        context.set_part(OpenXmlPart(package, part_uri))

        def visitor(element: etree._Element, inner_context: ValidationContext) -> None:
            if element.tag not in ACTIVE_CONTENT_TAGS:
                return

            element_name = _local_name(element)
            inner_context.add_error(
                error_type=ValidationErrorType.SEMANTIC,
                description=f"Active content element '{element_name}' found",
                node=element_name,
                severity=ValidationSeverity.ERROR,
                error_id="Sec_ActiveContentElement",
            )

        _walk_elements(root, context, visitor)
        if context.should_stop:
            return
