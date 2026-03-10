"""ODF semantic-core validator and rule registry."""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import (
    DRAW_NS,
    META_NS,
    NUMBER_NS,
    OFFICE_NS,
    PRESENTATION_NS,
    STYLE_NS,
    SVG_NS,
    TABLE_NS,
    TEXT_NS,
    XLINK_NS,
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
    # --- core ---
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
        id="ODFSEM005",
        family="core",
        description="meta.xml must contain office:meta element.",
    ),
    OdfSemanticRule(
        id="ODFSEM006",
        family="core",
        description="settings.xml must contain office:settings element.",
    ),
    # --- manifest ---
    OdfSemanticRule(
        id="ODFSEMMAN001",
        family="manifest",
        description="Key XML manifest entries must declare text/xml media type.",
    ),
    # --- style ---
    OdfSemanticRule(
        id="ODFSEMSTYLE001",
        family="style",
        description="Font face declarations must have svg:font-family attribute.",
    ),
    OdfSemanticRule(
        id="ODFSEMSTYLE002",
        family="style",
        description="Parent style references must resolve to defined styles.",
    ),
    OdfSemanticRule(
        id="ODFSEMSTYLE003",
        family="style",
        description="Data style references (style:data-style-name) must resolve.",
    ),
    OdfSemanticRule(
        id="ODFSEMSTYLE004",
        family="style",
        description="List style references must resolve to defined list styles.",
    ),
    OdfSemanticRule(
        id="ODFSEMSTYLE005",
        family="style",
        description="Page layout references in master pages must resolve.",
    ),
    # --- text ---
    OdfSemanticRule(
        id="ODFSEMTXT001",
        family="text",
        description="Text style references require styles.xml companion part.",
    ),
    OdfSemanticRule(
        id="ODFSEMTXT002",
        family="text",
        description="List level style references must resolve.",
    ),
    OdfSemanticRule(
        id="ODFSEMTXT003",
        family="text",
        description="Bookmark references must resolve to defined bookmarks.",
    ),
    # --- spreadsheet ---
    OdfSemanticRule(
        id="ODFSEMSS001",
        family="spreadsheet",
        description="Spreadsheet table names must be present and unique.",
    ),
    OdfSemanticRule(
        id="ODFSEMSS002",
        family="spreadsheet",
        description="Named ranges must reference valid table names.",
    ),
    OdfSemanticRule(
        id="ODFSEMSS003",
        family="spreadsheet",
        description="Column count in rows must not exceed table column definition.",
    ),
    # --- presentation ---
    OdfSemanticRule(
        id="ODFSEMPRES001",
        family="presentation",
        description="Presentation page names must be present and unique.",
    ),
    OdfSemanticRule(
        id="ODFSEMPRES002",
        family="presentation",
        description="Presentation must contain at least one draw:page.",
    ),
    OdfSemanticRule(
        id="ODFSEMPRES003",
        family="presentation",
        description="Presentation page layout references must resolve.",
    ),
    # --- reference (cross-part) ---
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
    OdfSemanticRule(
        id="ODFSEMREF003",
        family="reference",
        description="Font face declarations in content.xml must match styles.xml.",
    ),
    OdfSemanticRule(
        id="ODFSEMREF004",
        family="reference",
        description="Embedded object xlink:href references must resolve in package.",
    ),
    OdfSemanticRule(
        id="ODFSEMREF005",
        family="reference",
        description="Image xlink:href references must resolve in package.",
    ),
    # --- metadata ---
    OdfSemanticRule(
        id="ODFSEMMETA001",
        family="metadata",
        description="Document statistics attributes must be non-negative integers.",
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
        errors.extend(self._validate_meta_structure(parsed_parts))
        errors.extend(self._validate_settings_structure(parsed_parts))
        errors.extend(self._validate_manifest_key_media_types(package))
        errors.extend(self._validate_font_face_declarations(parsed_parts))
        errors.extend(self._validate_style_parent_refs(parsed_parts))
        errors.extend(self._validate_data_style_refs(parsed_parts))
        errors.extend(self._validate_list_style_refs(parsed_parts))
        errors.extend(self._validate_master_page_layouts(parsed_parts))
        errors.extend(self._validate_text_rules(package, parsed_parts))
        errors.extend(self._validate_text_list_levels(package, parsed_parts))
        errors.extend(self._validate_text_bookmarks(package, parsed_parts))
        errors.extend(self._validate_spreadsheet_rules(package, parsed_parts))
        errors.extend(self._validate_spreadsheet_named_ranges(package, parsed_parts))
        errors.extend(self._validate_spreadsheet_column_counts(package, parsed_parts))
        errors.extend(self._validate_presentation_rules(package, parsed_parts))
        errors.extend(self._validate_presentation_min_pages(package, parsed_parts))
        errors.extend(self._validate_presentation_page_layouts(package, parsed_parts))
        errors.extend(self._validate_cross_part_references(package, parsed_parts))
        errors.extend(self._validate_font_face_cross_part(package, parsed_parts))
        errors.extend(self._validate_embedded_object_refs(package, parsed_parts))
        errors.extend(self._validate_image_refs(package, parsed_parts))
        errors.extend(self._validate_meta_statistics(parsed_parts))
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

    # ── helpers ──────────────────────────────────────────────────────────

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

    def _manifest_entries_by_path(self, package: OdfPackage) -> dict[str, OdfManifestEntry]:
        entries: dict[str, OdfManifestEntry] = {}
        for entry in package.manifest:
            key = self._normalize_manifest_path(entry.full_path)
            if key and key not in entries:
                entries[key] = entry
        return entries

    @staticmethod
    def _collect_style_names(
        root: etree._Element,
        container_local: str = "automatic-styles",
    ) -> set[str]:
        """Collect style:name values from a given container."""
        names: set[str] = set()
        container = root.find(f"{{{OFFICE_NS}}}{container_local}")
        if container is None:
            return names
        for style in container:
            if not isinstance(style.tag, str):
                continue
            name = style.get(f"{{{STYLE_NS}}}name", "").strip()
            if name:
                names.add(name)
        return names

    @staticmethod
    def _collect_all_style_names(
        *roots: etree._Element | None,
    ) -> set[str]:
        """Collect all style:name values from automatic-styles and styles containers."""
        names: set[str] = set()
        for root in roots:
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    if name:
                        names.add(name)
        return names

    @staticmethod
    def _collect_list_style_names(*roots: etree._Element | None) -> set[str]:
        """Collect text:list-style names from automatic-styles and styles."""
        names: set[str] = set()
        for root in roots:
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for child in container:
                    if not isinstance(child.tag, str):
                        continue
                    qname = etree.QName(child)
                    if qname.namespace == TEXT_NS and qname.localname == "list-style":
                        name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                        if name:
                            names.add(name)
        return names

    @staticmethod
    def _collect_data_style_names(*roots: etree._Element | None) -> set[str]:
        """Collect number:* data style names."""
        names: set[str] = set()
        for root in roots:
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for child in container:
                    if not isinstance(child.tag, str):
                        continue
                    qname = etree.QName(child)
                    if qname.namespace == NUMBER_NS:
                        name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                        if name:
                            names.add(name)
        return names

    @staticmethod
    def _collect_page_layout_names(styles: etree._Element) -> set[str]:
        """Collect style:page-layout names from styles.xml."""
        names: set[str] = set()
        alm = styles.find(f"{{{OFFICE_NS}}}automatic-styles")
        if alm is None:
            return names
        for child in alm:
            if not isinstance(child.tag, str):
                continue
            qname = etree.QName(child)
            if qname.namespace == STYLE_NS and qname.localname == "page-layout":
                name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                if name:
                    names.add(name)
        return names

    @staticmethod
    def _collect_font_face_names(root: etree._Element) -> set[str]:
        """Collect font face declaration names from a root element."""
        names: set[str] = set()
        decls = root.find(f"{{{OFFICE_NS}}}font-face-decls")
        if decls is None:
            return names
        for child in decls:
            if not isinstance(child.tag, str):
                continue
            name = child.get(f"{{{STYLE_NS}}}name", "").strip()
            if name:
                names.add(name)
        return names

    # ── core rules ───────────────────────────────────────────────────────

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

    def _validate_meta_structure(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        meta = parsed_parts.get("meta.xml")
        if meta is None:
            return []
        if meta.find(f"{{{OFFICE_NS}}}meta") is not None:
            return []
        return [
            self._error(
                rule_id="ODFSEM005",
                error_type=ValidationErrorType.SCHEMA,
                description="meta.xml is missing required office:meta element",
                part_uri="/meta.xml",
            )
        ]

    def _validate_settings_structure(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        settings = parsed_parts.get("settings.xml")
        if settings is None:
            return []
        if settings.find(f"{{{OFFICE_NS}}}settings") is not None:
            return []
        return [
            self._error(
                rule_id="ODFSEM006",
                error_type=ValidationErrorType.SCHEMA,
                description="settings.xml is missing required office:settings element",
                part_uri="/settings.xml",
            )
        ]

    # ── manifest rules ───────────────────────────────────────────────────

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

    # ── style rules ──────────────────────────────────────────────────────

    def _validate_font_face_declarations(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSTYLE001: font-face-decl must have svg:font-family."""
        errors: list[ValidationError] = []
        for part_name in ("content.xml", "styles.xml"):
            root = parsed_parts.get(part_name)
            if root is None:
                continue
            decls = root.find(f"{{{OFFICE_NS}}}font-face-decls")
            if decls is None:
                continue
            for child in decls:
                if not isinstance(child.tag, str):
                    continue
                qname = etree.QName(child)
                if qname.namespace != STYLE_NS or qname.localname != "font-face":
                    continue
                font_family = child.get(f"{{{SVG_NS}}}font-family", "").strip()
                style_name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                if not font_family:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE001",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Font face declaration '{style_name or '(unnamed)'}' in "
                                f"{part_name} is missing required svg:font-family"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors

    def _validate_style_parent_refs(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSTYLE002: style:parent-style-name must resolve."""
        errors: list[ValidationError] = []
        content = parsed_parts.get("content.xml")
        styles = parsed_parts.get("styles.xml")
        all_names = self._collect_all_style_names(content, styles)

        for part_name in ("content.xml", "styles.xml"):
            root = parsed_parts.get(part_name)
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    parent = style.get(f"{{{STYLE_NS}}}parent-style-name", "").strip()
                    if not parent:
                        continue
                    if parent in all_names:
                        continue
                    style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE002",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Style '{style_name or '(unnamed)'}' in {part_name} "
                                f"references parent style '{parent}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors

    def _validate_data_style_refs(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSTYLE003: style:data-style-name must resolve to a number:* style."""
        errors: list[ValidationError] = []
        content = parsed_parts.get("content.xml")
        styles = parsed_parts.get("styles.xml")
        data_names = self._collect_data_style_names(content, styles)

        for part_name in ("content.xml", "styles.xml"):
            root = parsed_parts.get(part_name)
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    ref = style.get(f"{{{STYLE_NS}}}data-style-name", "").strip()
                    if not ref:
                        continue
                    if ref in data_names:
                        continue
                    style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE003",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Style '{style_name or '(unnamed)'}' in {part_name} "
                                f"references data style '{ref}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors

    def _validate_list_style_refs(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSTYLE004: style:list-style-name must resolve."""
        errors: list[ValidationError] = []
        content = parsed_parts.get("content.xml")
        styles = parsed_parts.get("styles.xml")
        list_names = self._collect_list_style_names(content, styles)
        if not list_names and content is None and styles is None:
            return errors

        for part_name in ("content.xml", "styles.xml"):
            root = parsed_parts.get(part_name)
            if root is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = root.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for style in container:
                    if not isinstance(style.tag, str):
                        continue
                    ref = style.get(f"{{{STYLE_NS}}}list-style-name", "").strip()
                    if not ref:
                        continue
                    if ref in list_names:
                        continue
                    style_name = style.get(f"{{{STYLE_NS}}}name", "").strip()
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSTYLE004",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Style '{style_name or '(unnamed)'}' in {part_name} "
                                f"references list style '{ref}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors

    def _validate_master_page_layouts(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSTYLE005: master-page page-layout-name must resolve."""
        errors: list[ValidationError] = []
        styles = parsed_parts.get("styles.xml")
        if styles is None:
            return errors

        layout_names = self._collect_page_layout_names(styles)
        master_pages = styles.find(f"{{{OFFICE_NS}}}master-styles")
        if master_pages is None:
            return errors

        for mp in master_pages:
            if not isinstance(mp.tag, str):
                continue
            qname = etree.QName(mp)
            if qname.namespace != STYLE_NS or qname.localname != "master-page":
                continue
            layout_ref = mp.get(f"{{{STYLE_NS}}}page-layout-name", "").strip()
            if not layout_ref:
                continue
            if layout_ref in layout_names:
                continue
            mp_name = mp.get(f"{{{STYLE_NS}}}name", "").strip()
            errors.append(
                self._error(
                    rule_id="ODFSEMSTYLE005",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Master page '{mp_name or '(unnamed)'}' references "
                        f"page layout '{layout_ref}' which is not defined"
                    ),
                    part_uri="/styles.xml",
                )
            )
        return errors

    # ── text rules ───────────────────────────────────────────────────────

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

    def _validate_text_list_levels(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMTXT002: text:list must reference defined list styles."""
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        styles = parsed_parts.get("styles.xml")
        list_names = self._collect_list_style_names(content, styles)

        reported: set[str] = set()
        for lst in content.iter(f"{{{TEXT_NS}}}list"):
            ref = lst.get(f"{{{TEXT_NS}}}style-name", "").strip()
            if not ref or ref in list_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT002",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"text:list references list style '{ref}' which is not defined"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors

    def _validate_text_bookmarks(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMTXT003: bookmark-ref must point to defined bookmarks."""
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        bookmark_names: set[str] = set()
        for elem in content.iter(f"{{{TEXT_NS}}}bookmark", f"{{{TEXT_NS}}}bookmark-start"):
            name = elem.get(f"{{{TEXT_NS}}}name", "").strip()
            if name:
                bookmark_names.add(name)

        reported: set[str] = set()
        for ref_elem in content.iter(f"{{{TEXT_NS}}}bookmark-ref"):
            ref_name = ref_elem.get(f"{{{TEXT_NS}}}ref-name", "").strip()
            if not ref_name or ref_name in bookmark_names or ref_name in reported:
                continue
            reported.add(ref_name)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Bookmark reference '{ref_name}' does not resolve to a "
                        "defined bookmark"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors

    # ── spreadsheet rules ────────────────────────────────────────────────

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

    def _validate_spreadsheet_named_ranges(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSS002: named range table references must resolve."""
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        table_names: set[str] = set()
        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            if name:
                table_names.add(name)

        for nr in content.xpath(
            ".//table:named-range", namespaces={"table": TABLE_NS}
        ):
            base_cell = nr.get(f"{{{TABLE_NS}}}base-cell-address", "").strip()
            if not base_cell:
                continue
            # base-cell-address format: $TableName.$A$1 or $'Table Name'.$A$1
            table_ref = base_cell.split(".")[0].strip("$").strip("'")
            if table_ref and table_ref not in table_names:
                range_name = nr.get(f"{{{TABLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS002",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Named range '{range_name or '(unnamed)'}' references "
                            f"table '{table_ref}' which does not exist"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors

    def _validate_spreadsheet_column_counts(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMSS003: row cell count must not exceed declared column count."""
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            table_name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            col_count = 0
            for col in table.iterchildren(f"{{{TABLE_NS}}}table-column"):
                repeat = col.get(
                    f"{{{TABLE_NS}}}number-columns-repeated", "1"
                ).strip()
                try:
                    col_count += int(repeat)
                except ValueError:
                    col_count += 1

            if col_count == 0:
                continue

            for row in table.iterchildren(f"{{{TABLE_NS}}}table-row"):
                cell_count = 0
                for cell in row.iterchildren(
                    f"{{{TABLE_NS}}}table-cell",
                    f"{{{TABLE_NS}}}covered-table-cell",
                ):
                    repeat = cell.get(
                        f"{{{TABLE_NS}}}number-columns-repeated", "1"
                    ).strip()
                    try:
                        cell_count += int(repeat)
                    except ValueError:
                        cell_count += 1

                if cell_count > col_count:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSS003",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Table '{table_name or '(unnamed)'}' has a row with "
                                f"{cell_count} cells but only {col_count} columns defined"
                            ),
                            part_uri="/content.xml",
                            severity=ValidationSeverity.WARNING,
                        )
                    )
                    break  # one warning per table is enough
        return errors

    # ── presentation rules ───────────────────────────────────────────────

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

    def _validate_presentation_min_pages(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMPRES002: presentation must have at least one draw:page."""
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return []
        content = parsed_parts.get("content.xml")
        if content is None:
            return []

        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            return []
        pres = body.find(f"{{{OFFICE_NS}}}presentation")
        if pres is None:
            return []

        pages = pres.findall(f"{{{DRAW_NS}}}page")
        if pages:
            return []

        return [
            self._error(
                rule_id="ODFSEMPRES002",
                error_type=ValidationErrorType.SEMANTIC,
                description="Presentation body contains no draw:page elements",
                part_uri="/content.xml",
            )
        ]

    def _validate_presentation_page_layouts(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMPRES003: presentation-page-layout-name must resolve."""
        errors: list[ValidationError] = []
        mimetype = package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.presentation"):
            return errors
        content = parsed_parts.get("content.xml")
        styles = parsed_parts.get("styles.xml")
        if content is None:
            return errors

        # Collect presentation:presentation-page-layout names
        layout_names: set[str] = set()
        for source in (content, styles):
            if source is None:
                continue
            for container_local in ("automatic-styles", "styles"):
                container = source.find(f"{{{OFFICE_NS}}}{container_local}")
                if container is None:
                    continue
                for child in container:
                    if not isinstance(child.tag, str):
                        continue
                    qname = etree.QName(child)
                    if (
                        qname.namespace == STYLE_NS
                        and qname.localname == "presentation-page-layout"
                    ):
                        name = child.get(f"{{{STYLE_NS}}}name", "").strip()
                        if name:
                            layout_names.add(name)

        if not layout_names:
            return errors

        reported: set[str] = set()
        for page in content.xpath(".//draw:page", namespaces={"draw": DRAW_NS}):
            ref = page.get(f"{{{PRESENTATION_NS}}}presentation-page-layout-name", "").strip()
            if not ref or ref in layout_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMPRES003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Presentation page layout '{ref}' is not defined"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors

    # ── cross-part reference rules ───────────────────────────────────────

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

    def _validate_font_face_cross_part(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMREF003: font faces in content.xml should exist in styles.xml too."""
        errors: list[ValidationError] = []
        content = parsed_parts.get("content.xml")
        styles = parsed_parts.get("styles.xml")
        if content is None or styles is None:
            return errors
        if "styles.xml" not in package.manifest_paths():
            return errors

        content_fonts = self._collect_font_face_names(content)
        styles_fonts = self._collect_font_face_names(styles)

        for font in sorted(content_fonts - styles_fonts):
            errors.append(
                self._error(
                    rule_id="ODFSEMREF003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Font face '{font}' declared in content.xml but not in styles.xml"
                    ),
                    part_uri="/content.xml",
                    severity=ValidationSeverity.WARNING,
                )
            )
        return errors

    def _validate_embedded_object_refs(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMREF004: embedded object xlink:href must resolve in package."""
        errors: list[ValidationError] = []
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        zip_members = package._zip_members()
        manifest_paths = package.manifest_paths()
        reported: set[str] = set()

        for obj in content.iter(f"{{{DRAW_NS}}}object", f"{{{DRAW_NS}}}object-ole"):
            href = obj.get(f"{{{XLINK_NS}}}href", "").strip()
            if not href or href.startswith("http://") or href.startswith("https://"):
                continue
            if href.startswith("./"):
                href = href[2:]
            if href in reported:
                continue

            # Embedded objects may be directories (with content.xml inside) or files
            found = (
                href in zip_members
                or href.rstrip("/") + "/content.xml" in zip_members
                or href in manifest_paths
                or href.rstrip("/") + "/" in manifest_paths
            )
            if not found:
                reported.add(href)
                errors.append(
                    self._error(
                        rule_id="ODFSEMREF004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Embedded object href '{href}' does not resolve in package"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors

    def _validate_image_refs(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMREF005: image xlink:href must resolve in package."""
        errors: list[ValidationError] = []
        content = parsed_parts.get("content.xml")
        if content is None:
            return errors

        zip_members = package._zip_members()
        reported: set[str] = set()

        for img in content.iter(f"{{{DRAW_NS}}}image"):
            href = img.get(f"{{{XLINK_NS}}}href", "").strip()
            if not href or href.startswith("http://") or href.startswith("https://"):
                continue
            if href.startswith("./"):
                href = href[2:]
            if href in reported:
                continue
            if href not in zip_members:
                reported.add(href)
                errors.append(
                    self._error(
                        rule_id="ODFSEMREF005",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Image href '{href}' does not resolve in package"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors

    # ── metadata rules ───────────────────────────────────────────────────

    def _validate_meta_statistics(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        """ODFSEMMETA001: document-statistic attributes must be non-negative integers."""
        errors: list[ValidationError] = []
        meta = parsed_parts.get("meta.xml")
        if meta is None:
            return errors

        office_meta = meta.find(f"{{{OFFICE_NS}}}meta")
        if office_meta is None:
            return errors

        stats = office_meta.find(f"{{{META_NS}}}document-statistic")
        if stats is None:
            return errors

        count_attrs = [
            f"{{{META_NS}}}page-count",
            f"{{{META_NS}}}table-count",
            f"{{{META_NS}}}image-count",
            f"{{{META_NS}}}object-count",
            f"{{{META_NS}}}paragraph-count",
            f"{{{META_NS}}}word-count",
            f"{{{META_NS}}}character-count",
            f"{{{META_NS}}}non-whitespace-character-count",
            f"{{{META_NS}}}cell-count",
            f"{{{META_NS}}}sentence-count",
            f"{{{META_NS}}}syllable-count",
            f"{{{META_NS}}}row-count",
        ]

        for attr in count_attrs:
            value = stats.get(attr)
            if value is None:
                continue
            value = value.strip()
            try:
                n = int(value)
                if n < 0:
                    raise ValueError("negative")
            except ValueError:
                local = attr.split("}")[-1]
                errors.append(
                    self._error(
                        rule_id="ODFSEMMETA001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Document statistic '{local}' has invalid value "
                            f"'{value}' (expected non-negative integer)"
                        ),
                        part_uri="/meta.xml",
                    )
                )
        return errors
