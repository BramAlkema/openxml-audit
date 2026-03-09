"""ODF semantic-core validator and rule registry."""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import (
    DRAW_NS,
    OFFICE_NS,
    STYLE_NS,
    TABLE_NS,
    TEXT_NS,
    normalize_manifest_path,
    normalize_part_uri,
)
from openxml_audit.odf.package import OdfManifestEntry, OdfPackage


@dataclass(frozen=True)
class OdfSemanticRule:
    """Stable semantic-core rule metadata."""

    id: str
    family: str
    description: str


RULES: tuple[OdfSemanticRule, ...] = (
    OdfSemanticRule(
        id="ODFSEM001",
        family="core",
        description="Core XML parts must use expected office:* root elements.",
    ),
    OdfSemanticRule(
        id="ODFSEM002",
        family="core",
        description="content.xml must contain office:body.",
    ),
    OdfSemanticRule(
        id="ODFSEM003",
        family="core",
        description="content.xml office:body must contain exactly one document-class element.",
    ),
    OdfSemanticRule(
        id="ODFSEM004",
        family="core",
        description="content.xml body class must match package mimetype.",
    ),
    OdfSemanticRule(
        id="ODFSEMMAN001",
        family="manifest",
        description="Key XML manifest entries must declare text/xml media type.",
    ),
    OdfSemanticRule(
        id="ODFSEMTXT001",
        family="text",
        description="Text style references require styles.xml companion part.",
    ),
    OdfSemanticRule(
        id="ODFSEMSS001",
        family="spreadsheet",
        description="Spreadsheet table names must be present and unique.",
    ),
    OdfSemanticRule(
        id="ODFSEMPRES001",
        family="presentation",
        description="Presentation page names must be present and unique.",
    ),
    OdfSemanticRule(
        id="ODFSEMREF001",
        family="reference",
        description="Presentation master-page references require styles.xml companion part.",
    ),
    OdfSemanticRule(
        id="ODFSEMREF002",
        family="reference",
        description="Presentation master-page references must resolve in styles.xml.",
    ),
)


def get_odf_semantic_rules() -> tuple[OdfSemanticRule, ...]:
    """Return semantic-core rule metadata with stable identifiers."""
    return RULES


class OdfSemanticValidator:
    """Semantic-core checks for ODF packages."""

    CORE_ROOTS = {
        "content.xml": "document-content",
        "styles.xml": "document-styles",
        "meta.xml": "document-meta",
        "settings.xml": "document-settings",
    }
    CONTENT_BODY_BY_MIMETYPE = (
        ("application/vnd.oasis.opendocument.text", "text"),
        ("application/vnd.oasis.opendocument.spreadsheet", "spreadsheet"),
        ("application/vnd.oasis.opendocument.presentation", "presentation"),
    )
    KEY_XML_MEDIA_TYPE_PARTS = (
        "content.xml",
        "styles.xml",
        "meta.xml",
        "settings.xml",
        "META-INF/documentsignatures.xml",
    )

    def validate(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        errors.extend(self._validate_core_roots(package, parsed_parts))
        errors.extend(self._validate_body_document_class(package, parsed_parts))
        errors.extend(self._validate_manifest_key_media_types(package))
        errors.extend(self._validate_text_rules(package, parsed_parts))
        errors.extend(self._validate_spreadsheet_rules(package, parsed_parts))
        errors.extend(self._validate_presentation_rules(package, parsed_parts))
        errors.extend(self._validate_cross_part_references(package, parsed_parts))
        return errors

    @staticmethod
    def _normalize_part_uri(part_path: str) -> str:
        return normalize_part_uri(part_path)

    @staticmethod
    def _normalize_manifest_path(path: str) -> str:
        return normalize_manifest_path(path)

    @staticmethod
    def _error(
        *,
        rule_id: str,
        error_type: ValidationErrorType,
        description: str,
        part_uri: str,
        severity: ValidationSeverity = ValidationSeverity.ERROR,
    ) -> ValidationError:
        return ValidationError(
            error_type=error_type,
            description=description,
            part_uri=part_uri,
            severity=severity,
            id=rule_id,
        )

    def _validate_core_roots(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        manifest_paths = package.manifest_paths()
        for part, expected_root in self.CORE_ROOTS.items():
            if part not in manifest_paths:
                continue
            root = parsed_parts.get(part)
            if root is None:
                continue
            qname = etree.QName(root)
            if qname.namespace == OFFICE_NS and qname.localname == expected_root:
                continue
            errors.append(
                self._error(
                    rule_id="ODFSEM001",
                    error_type=ValidationErrorType.SCHEMA,
                    description=(
                        f"{part} root element must be office:{expected_root} "
                        f"(found '{root.tag}')"
                    ),
                    part_uri=self._normalize_part_uri(part),
                )
            )
        return errors

    @classmethod
    def _expected_content_body_local(cls, mimetype: str | None) -> str | None:
        if not mimetype:
            return None
        for prefix, local in cls.CONTENT_BODY_BY_MIMETYPE:
            if mimetype.startswith(prefix):
                return local
        return None

    @staticmethod
    def _element_children(element: etree._Element) -> list[etree._Element]:
        return [child for child in element if isinstance(child.tag, str)]

    def _validate_body_document_class(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        if "content.xml" not in package.manifest_paths():
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            errors.append(
                self._error(
                    rule_id="ODFSEM002",
                    error_type=ValidationErrorType.SCHEMA,
                    description="content.xml is missing required office:body element",
                    part_uri="/content.xml",
                )
            )
            return errors

        body_children = self._element_children(body)
        if len(body_children) != 1:
            errors.append(
                self._error(
                    rule_id="ODFSEM003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        "content.xml office:body must contain exactly one document type "
                        f"element (found {len(body_children)})"
                    ),
                    part_uri="/content.xml",
                )
            )
            if not body_children:
                return errors

        expected_body_local = self._expected_content_body_local(package.mimetype)
        if expected_body_local is None:
            return errors

        first = body_children[0]
        first_qname = etree.QName(first)
        if first_qname.namespace == OFFICE_NS and first_qname.localname == expected_body_local:
            return errors

        errors.append(
            self._error(
                rule_id="ODFSEM004",
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    "content.xml body type does not match mimetype "
                    f"(expected office:{expected_body_local}, found '{first.tag}')"
                ),
                part_uri="/content.xml",
            )
        )
        return errors

    def _manifest_entries_by_path(self, package: OdfPackage) -> dict[str, OdfManifestEntry]:
        entries: dict[str, OdfManifestEntry] = {}
        for entry in package.manifest:
            key = self._normalize_manifest_path(entry.full_path)
            if key and key not in entries:
                entries[key] = entry
        return entries

    def _validate_manifest_key_media_types(self, package: OdfPackage) -> list[ValidationError]:
        errors: list[ValidationError] = []
        entries = self._manifest_entries_by_path(package)
        for path in self.KEY_XML_MEDIA_TYPE_PARTS:
            entry = entries.get(path)
            if entry is None:
                continue
            media_type = entry.media_type.strip().lower()
            if media_type == "text/xml":
                continue
            errors.append(
                self._error(
                    rule_id="ODFSEMMAN001",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Manifest media-type for '{path}' should be 'text/xml' "
                        f"(found '{entry.media_type}')"
                    ),
                    part_uri="/META-INF/manifest.xml",
                )
            )
        return errors

    @staticmethod
    def _has_text_style_references(content: etree._Element) -> bool:
        style_attr = f"{{{TEXT_NS}}}style-name"
        for elem in content.iter():
            tag = elem.tag
            if not isinstance(tag, str) or not tag.startswith(f"{{{TEXT_NS}}}"):
                continue
            value = elem.get(style_attr)
            if value is not None and value.strip():
                return True
        return False

    def _validate_text_rules(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        if "styles.xml" in package.manifest_paths():
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        if not self._has_text_style_references(content):
            return errors

        errors.append(
            self._error(
                rule_id="ODFSEMTXT001",
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    "content.xml contains text:style-name references but styles.xml "
                    "is not declared in manifest.xml"
                ),
                part_uri="/content.xml",
            )
        )
        return errors

    def _validate_spreadsheet_rules(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            if not name:
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="Spreadsheet table is missing required table:name",
                        part_uri="/content.xml",
                    )
                )
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate spreadsheet table name '{name}'",
                        part_uri="/content.xml",
                    )
                )
                continue
            seen.add(name)
        return errors

    def _validate_presentation_rules(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            name = page.get(f"{{{DRAW_NS}}}name", "").strip()
            if not name:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="Presentation page is missing required draw:name",
                        part_uri="/content.xml",
                    )
                )
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMPRES001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate presentation page name '{name}'",
                        part_uri="/content.xml",
                    )
                )
                continue
            seen.add(name)
        return errors

    @staticmethod
    def _collect_presentation_master_refs(content: etree._Element) -> set[str]:
        refs: set[str] = set()
        pages = content.xpath(
            ".//draw:page[@draw:master-page-name]",
            namespaces={"draw": DRAW_NS},
        )
        for page in pages:
            value = page.get(f"{{{DRAW_NS}}}master-page-name", "").strip()
            if value:
                refs.add(value)
        return refs

    @staticmethod
    def _collect_master_page_definitions(styles: etree._Element) -> set[str]:
        names: set[str] = set()
        for page in styles.xpath(
            ".//style:master-page[@style:name]",
            namespaces={"style": STYLE_NS},
        ):
            name = page.get(f"{{{STYLE_NS}}}name", "").strip()
            if name:
                names.add(name)
        return names

    def _validate_cross_part_references(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        refs = self._collect_presentation_master_refs(content)
        if not refs:
            return errors
        if "styles.xml" not in package.manifest_paths():
            errors.append(
                self._error(
                    rule_id="ODFSEMREF001",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        "draw:master-page-name references require styles.xml "
                        "to be declared in manifest.xml"
                    ),
                    part_uri="/content.xml",
                )
            )
            return errors

        styles = parsed_parts.get("styles.xml")
        if styles is None:
            return errors
        definitions = self._collect_master_page_definitions(styles)
        for ref in sorted(refs):
            if ref in definitions:
                continue
            errors.append(
                self._error(
                    rule_id="ODFSEMREF002",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"draw:master-page-name '{ref}' does not resolve to a "
                        "style:master-page in styles.xml"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors
