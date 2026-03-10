"""Core ODF constraints: root elements, body structure, meta, settings."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType
from openxml_audit.odf._helpers import OFFICE_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule


class CoreRootConstraint(OdfConstraint):
    """ODFSEM001: Core XML parts must use expected office:* root elements."""

    CORE_ROOTS = {
        "content.xml": "document-content",
        "styles.xml": "document-styles",
        "meta.xml": "document-meta",
        "settings.xml": "document-settings",
    }

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEM001",
            family="core",
            description="Core XML parts must use expected office:* root elements.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        manifest_paths = ctx.package.manifest_paths()
        for part, expected_root in self.CORE_ROOTS.items():
            if part not in manifest_paths:
                continue
            root = ctx.parsed_parts.get(part)
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


class BodyDocumentClassConstraint(OdfConstraint):
    """ODFSEM002/003/004: content.xml body structure and mimetype match."""

    CONTENT_BODY_BY_MIMETYPE = (
        ("application/vnd.oasis.opendocument.text", "text"),
        ("application/vnd.oasis.opendocument.spreadsheet", "spreadsheet"),
        ("application/vnd.oasis.opendocument.presentation", "presentation"),
    )

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEM002",
            family="core",
            description="content.xml must contain office:body.",
        )

    @staticmethod
    def _element_children(element: etree._Element) -> list[etree._Element]:
        return [child for child in element if isinstance(child.tag, str)]

    @classmethod
    def _expected_content_body_local(cls, mimetype: str | None) -> str | None:
        if not mimetype:
            return None
        for prefix, local in cls.CONTENT_BODY_BY_MIMETYPE:
            if mimetype.startswith(prefix):
                return local
        return None

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        if "content.xml" not in ctx.package.manifest_paths():
            return errors
        content = ctx.parsed_parts.get("content.xml")
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

        expected_body_local = self._expected_content_body_local(ctx.package.mimetype)
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


class MetaStructureConstraint(OdfConstraint):
    """ODFSEM005: meta.xml must contain office:meta element."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEM005",
            family="core",
            description="meta.xml must contain office:meta element.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        meta = ctx.parsed_parts.get("meta.xml")
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


class SettingsStructureConstraint(OdfConstraint):
    """ODFSEM006: settings.xml must contain office:settings element."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEM006",
            family="core",
            description="settings.xml must contain office:settings element.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        settings = ctx.parsed_parts.get("settings.xml")
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
