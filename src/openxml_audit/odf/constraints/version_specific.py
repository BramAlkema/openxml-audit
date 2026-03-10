"""Version-specific ODF validation constraints (M4).

These rules enforce version-dependent semantics: features that only
exist in certain ODF versions, and cross-version consistency checks.
"""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import (
    DRAW_NS,
    OFFICE_NS,
    TABLE_NS,
    TEXT_NS,
)
from openxml_audit.odf.constraints.base import (
    EvaluationContext,
    OdfConstraint,
    OdfSemanticRule,
    parse_version_tuple,
)


class VersionAttributePresentConstraint(OdfConstraint):
    """ODFSEMVER001: office:version should be present on document root."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER001",
            family="version",
            description="Document root elements should declare office:version.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part in ("content.xml", "styles.xml"):
            root = ctx.parsed_parts.get(part)
            if root is None:
                continue
            version = root.get(f"{{{OFFICE_NS}}}version", "").strip()
            if not version:
                errors.append(
                    self._error(
                        rule_id="ODFSEMVER001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"{part} root element is missing office:version attribute"
                        ),
                        part_uri=self._normalize_part_uri(part),
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors


class VersionConsistencyConstraint(OdfConstraint):
    """ODFSEMVER002: office:version across parts and manifest must be consistent."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER002",
            family="version",
            description="ODF version must be consistent across parts and manifest.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        versions: dict[str, str] = {}

        for part in ("content.xml", "styles.xml", "meta.xml", "settings.xml"):
            root = ctx.parsed_parts.get(part)
            if root is None:
                continue
            version = root.get(f"{{{OFFICE_NS}}}version", "").strip()
            if version:
                versions[part] = version

        manifest_version = ctx.package.manifest_version
        if manifest_version:
            versions["manifest.xml"] = manifest_version.strip()

        if len(set(versions.values())) <= 1:
            return errors

        unique = sorted(set(versions.values()))
        parts_detail = ", ".join(
            f"{part}={ver}" for part, ver in sorted(versions.items())
        )
        errors.append(
            self._error(
                rule_id="ODFSEMVER002",
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    f"Inconsistent ODF versions detected ({', '.join(unique)}): "
                    f"{parts_detail}"
                ),
                part_uri="/content.xml",
                severity=ValidationSeverity.WARNING,
            )
        )
        return errors


class RdfMetadataVersionConstraint(OdfConstraint):
    """ODFSEMVER003: RDF metadata (manifest.rdf) requires ODF 1.2+."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER003",
            family="version",
            description="RDF metadata files require ODF 1.2 or later.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []

        # Check if manifest declares an RDF file
        has_rdf = False
        for entry in ctx.package.manifest:
            if entry.media_type == "application/rdf+xml":
                has_rdf = True
                break

        if not has_rdf:
            return errors

        doc_ver = parse_version_tuple(ctx.odf_version)
        if doc_ver is not None and doc_ver < (1, 2):
            errors.append(
                self._error(
                    rule_id="ODFSEMVER003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"RDF metadata declared in manifest but document version "
                        f"is {ctx.odf_version} (requires ODF 1.2+)"
                    ),
                    part_uri="/META-INF/manifest.xml",
                )
            )
        return errors


class NamedExpressionsVersionConstraint(OdfConstraint):
    """ODFSEMVER004: Document-level named-expressions require ODF 1.2+."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER004",
            family="version",
            description="Document-level named-expressions require ODF 1.2+.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        doc_ver = parse_version_tuple(ctx.odf_version)
        if doc_ver is None or doc_ver >= (1, 2):
            return errors

        # Named expressions as direct child of office:spreadsheet (document-level)
        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            return errors
        spreadsheet = body.find(f"{{{OFFICE_NS}}}spreadsheet")
        if spreadsheet is None:
            return errors

        for child in spreadsheet:
            if not isinstance(child.tag, str):
                continue
            qname = etree.QName(child)
            if (
                qname.namespace == TABLE_NS
                and qname.localname == "named-expressions"
            ):
                errors.append(
                    self._error(
                        rule_id="ODFSEMVER004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            "Document-level table:named-expressions found but "
                            f"document version is {ctx.odf_version} (requires ODF 1.2+)"
                        ),
                        part_uri="/content.xml",
                    )
                )
                break
        return errors


class DigitalSignatureVersionConstraint(OdfConstraint):
    """ODFSEMVER005: Digital signature entries require ODF 1.2+."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER005",
            family="version",
            description="Digital signature manifest entries require ODF 1.2+.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []

        has_sig = False
        for entry in ctx.package.manifest:
            path = entry.full_path.lower()
            if "documentsignatures" in path or "macrosignatures" in path:
                has_sig = True
                break

        if not has_sig:
            return errors

        doc_ver = parse_version_tuple(ctx.odf_version)
        if doc_ver is not None and doc_ver < (1, 2):
            errors.append(
                self._error(
                    rule_id="ODFSEMVER005",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Digital signature entries found in manifest but document "
                        f"version is {ctx.odf_version} (requires ODF 1.2+)"
                    ),
                    part_uri="/META-INF/manifest.xml",
                )
            )
        return errors


class ChangeTrackingVersionConstraint(OdfConstraint):
    """ODFSEMVER006: text:tracked-changes requires ODF 1.2+."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER006",
            family="version",
            description="Change tracking (text:tracked-changes) requires ODF 1.2+.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        doc_ver = parse_version_tuple(ctx.odf_version)
        if doc_ver is None or doc_ver >= (1, 2):
            return errors

        for _elem in content.iter(f"{{{TEXT_NS}}}tracked-changes"):
            errors.append(
                self._error(
                    rule_id="ODFSEMVER006",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        "text:tracked-changes found but document version is "
                        f"{ctx.odf_version} (requires ODF 1.2+)"
                    ),
                    part_uri="/content.xml",
                )
            )
            break
        return errors


class DrawEnhancedGeometryVersionConstraint(OdfConstraint):
    """ODFSEMVER007: draw:enhanced-geometry extended attributes require ODF 1.3+."""

    # Attributes added in ODF 1.3 for enhanced geometry
    _V13_ATTRS = (
        f"{{{DRAW_NS}}}extrusion-allowed",
        f"{{{DRAW_NS}}}text-path-allowed",
        f"{{{DRAW_NS}}}concentric-gradient-fill-allowed",
    )

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER007",
            family="version",
            description="Extended draw:enhanced-geometry attributes require ODF 1.3+.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        doc_ver = parse_version_tuple(ctx.odf_version)
        if doc_ver is None or doc_ver >= (1, 3):
            return errors

        for elem in content.iter(f"{{{DRAW_NS}}}enhanced-geometry"):
            for attr in self._V13_ATTRS:
                if elem.get(attr) is not None:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMVER007",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"draw:enhanced-geometry uses attribute "
                                f"'{etree.QName(attr).localname}' which requires "
                                f"ODF 1.3+ (document version is {ctx.odf_version})"
                            ),
                            part_uri="/content.xml",
                        )
                    )
                    return errors
        return errors


class PresentationAnimationVersionConstraint(OdfConstraint):
    """ODFSEMVER008: Presentation animation iterate elements require ODF 1.3+."""

    SMIL_NS = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"
    ANIM_NS = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0"

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMVER008",
            family="version",
            description="Extended presentation animation elements require ODF 1.3+.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        doc_ver = parse_version_tuple(ctx.odf_version)
        if doc_ver is None or doc_ver >= (1, 3):
            return errors

        # anim:iterate was formalized in ODF 1.3
        for _elem in content.iter(f"{{{self.ANIM_NS}}}iterate"):
            errors.append(
                self._error(
                    rule_id="ODFSEMVER008",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        "anim:iterate element found but document version is "
                        f"{ctx.odf_version} (requires ODF 1.3+)"
                    ),
                    part_uri="/content.xml",
                )
            )
            break
        return errors
