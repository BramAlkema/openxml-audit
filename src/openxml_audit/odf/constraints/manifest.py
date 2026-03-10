"""Manifest media-type constraint."""

from __future__ import annotations

from openxml_audit.errors import ValidationError, ValidationErrorType
from openxml_audit.odf._helpers import normalize_manifest_path
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule
from openxml_audit.odf.package import OdfManifestEntry


class ManifestMediaTypeConstraint(OdfConstraint):
    """ODFSEMMAN001: Key XML manifest entries must declare text/xml media type."""

    KEY_XML_MEDIA_TYPE_PARTS = (
        "content.xml",
        "styles.xml",
        "meta.xml",
        "settings.xml",
        "META-INF/documentsignatures.xml",
    )

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMMAN001",
            family="manifest",
            description="Key XML manifest entries must declare text/xml media type.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        entries: dict[str, OdfManifestEntry] = {}
        for entry in ctx.package.manifest:
            key = normalize_manifest_path(entry.full_path)
            if key and key not in entries:
                entries[key] = entry

        for path in self.KEY_XML_MEDIA_TYPE_PARTS:
            matched = entries.get(path)
            if matched is None:
                continue
            media_type = matched.media_type.strip().lower()
            if media_type == "text/xml":
                continue
            errors.append(
                self._error(
                    rule_id="ODFSEMMAN001",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Manifest media-type for '{path}' should be 'text/xml' "
                        f"(found '{matched.media_type}')"
                    ),
                    part_uri="/META-INF/manifest.xml",
                )
            )
        return errors
