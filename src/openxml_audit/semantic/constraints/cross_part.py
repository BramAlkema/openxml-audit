"""Cross-part semantic constraints."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.semantic.attributes import SemanticConstraint

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.package import OpenXmlPackage


_MAIN_PART_ALIASES = {
    "WorkbookPart",
    "MainDocumentPart",
    "PresentationPart",
}


def _build_xpath(element_xpath: str) -> str:
    xpath = element_xpath.lstrip("/")
    return f"//{xpath}" if xpath else "//"


def _qualified_attr_name(local_name: str, namespace: str | None) -> str:
    return f"{{{namespace}}}{local_name}" if namespace else local_name


def _part_keywords(part_name: str) -> list[str]:
    name = part_name
    if name.endswith("Part"):
        name = name[:-4]

    tokens = re.findall(r"[A-Z][a-z0-9]*|[a-z0-9]+", name)
    if not tokens:
        return []

    keywords = [tokens[-1].lower()]
    keywords.append("".join(token.lower() for token in tokens))
    return list(dict.fromkeys(keywords))


def _match_parts_by_name(package: OpenXmlPackage, part_name: str) -> list[str]:
    keywords = _part_keywords(part_name)
    if not keywords:
        return []

    matches = []
    for part_uri in package.list_parts():
        uri_lower = part_uri.lower()
        if any(keyword in uri_lower for keyword in keywords):
            matches.append(part_uri)
    return matches


def _resolve_part_uris(
    part_path: str,
    context: ValidationContext,
    package: OpenXmlPackage,
) -> list[str]:
    if part_path == ".":
        if context.part is None:
            return []
        return [context.part.uri]

    if part_path == "..":
        if context.part is None:
            return []
        parent = context.part.uri.rsplit("/", 1)[0]
        return [part_uri for part_uri in package.list_parts() if part_uri.startswith(f"{parent}/")]

    if part_path.startswith("/") and package.has_part(part_path):
        return [part_path]

    if not part_path.startswith("/"):
        candidate = f"/{part_path}"
        if package.has_part(candidate):
            return [candidate]

    normalized = part_path.lstrip("/")
    if normalized in _MAIN_PART_ALIASES:
        main_uri = package.get_main_document_uri()
        return [main_uri] if main_uri else []

    segments = normalized.split("/") if normalized else []
    if segments:
        last_segment = segments[-1]
        matched = _match_parts_by_name(package, last_segment)
        if len(matched) == 1:
            return matched

    return []


def _cache_key(context: ValidationContext, part_path: str, part_uris: list[str]) -> tuple[int, str]:
    package_id = id(context.package)
    if part_uris:
        return (package_id, ",".join(sorted(part_uris)))
    if part_path == "." and context.part is not None:
        return (package_id, context.part.uri)
    return (package_id, part_path)


@dataclass
class CrossPartCountConstraint(SemanticConstraint):
    """Validates an attribute against a count from another part.

    Matches schematron rules like:
        @x:cm < count(document('Part:/WorkbookPart/CellMetadataPart')//x:cellMetadata/x:bk) + 1
    """

    attribute: str
    part_path: str
    element_xpath: str
    count_offset: int = 0
    namespace_map: dict[str, str] = field(default_factory=dict)
    attribute_namespace: str | None = None
    _count_cache: dict[tuple[int, str], int | None] = field(
        default_factory=dict, init=False, repr=False
    )
    _xpath: str = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self._xpath = _build_xpath(self.element_xpath)

    def validate(self, element: etree._Element, context: ValidationContext) -> bool:
        attr_name = _qualified_attr_name(self.attribute, self.attribute_namespace)
        if attr_name not in element.attrib:
            return True

        raw_value = element.attrib[attr_name]
        try:
            value = float(raw_value)
        except ValueError:
            context.add_semantic_error(
                f"Attribute '{self.attribute}' must be numeric",
                node=self.attribute,
            )
            return False

        count = self._get_count(context)
        if count is None:
            return True

        limit = count + self.count_offset
        if value >= limit:
            target = self._describe_target()
            context.add_semantic_error(
                f"Attribute '{self.attribute}' value {raw_value} must be < {limit} "
                f"(count {count} from {target})",
                node=self.attribute,
            )
            return False

        return True

    def _describe_target(self) -> str:
        if self.part_path == ".":
            part_desc = "current part"
        else:
            part_desc = f"Part:{self.part_path}"
        return f"{part_desc} xpath='{self._xpath}'"

    def _get_count(self, context: ValidationContext) -> int | None:
        package = context.package
        if package is None:
            return None

        part_uris = _resolve_part_uris(self.part_path, context, package)
        cache_key = _cache_key(context, self.part_path, part_uris)
        if cache_key in self._count_cache:
            return self._count_cache[cache_key]

        if part_uris:
            count = sum(self._count_in_part(package, uri) for uri in part_uris)
        else:
            count = self._count_by_xpath_scan(package)

        self._count_cache[cache_key] = count
        return count

    def _count_by_xpath_scan(self, package: OpenXmlPackage) -> int | None:
        total = 0
        found = False

        for part_uri in package.list_parts():
            count = self._count_in_part(package, part_uri)
            if count:
                found = True
                total += count

        return total if found else None

    def _count_in_part(self, package: OpenXmlPackage, part_uri: str) -> int:
        xml = package.get_part_xml(part_uri)
        if xml is None:
            return 0
        try:
            return len(xml.xpath(self._xpath, namespaces=self.namespace_map))
        except etree.XPathEvalError:
            return 0


@dataclass
class CrossPartReferenceConstraint(SemanticConstraint):
    """Validates that an attribute value exists in a referenced part XPath set."""

    attribute: str
    part_path: str
    element_xpath: str
    target_attribute: str
    namespace_map: dict[str, str] = field(default_factory=dict)
    attribute_namespace: str | None = None
    target_attribute_namespace: str | None = None
    _value_cache: dict[tuple[int, str], set[str] | None] = field(
        default_factory=dict, init=False, repr=False
    )
    _xpath: str = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self._xpath = _build_xpath(self.element_xpath)

    def validate(self, element: etree._Element, context: ValidationContext) -> bool:
        attr_name = _qualified_attr_name(self.attribute, self.attribute_namespace)
        if attr_name not in element.attrib:
            return True

        value = element.attrib[attr_name]
        valid_values = self._get_reference_values(context)
        if valid_values is None:
            return True
        if value in valid_values:
            return True

        context.add_semantic_error(
            f"Attribute '{self.attribute}' value '{value}' was not found in "
            f"Part:{self.part_path} xpath='{self._xpath}'/@{self.target_attribute}",
            node=self.attribute,
        )
        return False

    def _get_reference_values(self, context: ValidationContext) -> set[str] | None:
        package = context.package
        if package is None:
            return None

        part_uris = _resolve_part_uris(self.part_path, context, package)
        cache_key = _cache_key(context, self.part_path, part_uris)
        if cache_key in self._value_cache:
            return self._value_cache[cache_key]

        values: set[str] = set()
        if part_uris:
            for part_uri in part_uris:
                values.update(self._extract_values_from_part(package, part_uri))
        else:
            for part_uri in package.list_parts():
                values.update(self._extract_values_from_part(package, part_uri))

        cached = values if values else None
        self._value_cache[cache_key] = cached
        return cached

    def _extract_values_from_part(self, package: OpenXmlPackage, part_uri: str) -> set[str]:
        xml = package.get_part_xml(part_uri)
        if xml is None:
            return set()

        try:
            nodes = xml.xpath(self._xpath, namespaces=self.namespace_map)
        except etree.XPathEvalError:
            return set()

        attr_name = _qualified_attr_name(self.target_attribute, self.target_attribute_namespace)
        values: set[str] = set()
        for node in nodes:
            if not isinstance(node, etree._Element):
                continue
            value = node.get(attr_name)
            if value:
                values.add(value)
        return values
