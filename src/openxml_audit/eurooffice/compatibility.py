"""Version-bound classification of known EuroOffice editor drift.

This module does not change, remove, or downgrade validator findings.  It adds
an explicitly scoped compatibility view over a :class:`ValidationResult` so a
consumer can distinguish an observed EuroOffice serializer quirk from new
drift.  Profiles are deliberately exact: an unknown Document Server,
connector, validation format, finding family, or occurrence count is never
waived.

The first profile was calibrated from TokenOOXML documents saved through the
Nextcloud EuroOffice editor callback.  It covers static package/schema
findings only.  It does not claim that document semantics were preserved.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any
from zipfile import BadZipFile, ZipFile

from openxml_audit.errors import (
    FileFormat,
    SourceClass,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.parity_normalization import normalize_error_id

__all__ = [
    "TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID",
    "EuroOfficeCompatibilityReport",
    "EuroOfficeCompatibilityStatus",
    "classify_eurooffice_compatibility",
    "supported_eurooffice_compatibility_profiles",
]

TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID = "tokenooxml-eurooffice-editor-9.3.1.37-connector-11.0.1"

_SEMANTIC_PRESERVATION = "not-assessed"
_EVIDENCE_SCOPE = (
    "TokenOOXML documents saved through the Nextcloud EuroOffice editor callback; "
    "static OOXML findings only"
)

_OOXML_FORMATS = frozenset(
    {
        FileFormat.OFFICE_2007,
        FileFormat.MICROSOFT_365,
    }
)


class EuroOfficeCompatibilityStatus(str, Enum):
    """Outcome of applying a version-bound EuroOffice compatibility profile."""

    STRICT_CLEAN = "strict-clean"
    ACCEPTED_KNOWN_DRIFT = "accepted-known-drift"
    UNEXPECTED_DRIFT = "unexpected-drift"
    UNVERIFIED_ENVIRONMENT = "unverified-environment"


@dataclass(frozen=True, slots=True)
class _KnownDriftRule:
    rule_id: str
    extensions: frozenset[str]
    formats: frozenset[FileFormat]
    severity: ValidationSeverity
    error_type: ValidationErrorType
    source_class: SourceClass
    error_id: str
    part_uri: str
    node: str | None
    description: str
    max_occurrences: int
    paths: frozenset[str] | None = None
    path_pattern: re.Pattern[str] | None = None
    requires_no_bin_payload: bool = False

    def matches(
        self,
        finding: ValidationError,
        *,
        extension: str,
        file_format: FileFormat,
    ) -> bool:
        if extension not in self.extensions or file_format not in self.formats:
            return False
        if finding.severity is not self.severity:
            return False
        if finding.error_type is not self.error_type:
            return False
        if finding.source_class is not self.source_class:
            return False
        if normalize_error_id(finding) != self.error_id:
            return False
        if finding.part_uri != self.part_uri or finding.node != self.node:
            return False
        if finding.description != self.description:
            return False
        if self.paths is not None and finding.path not in self.paths:
            return False
        return self.path_pattern is None or self.path_pattern.fullmatch(finding.path) is not None


@dataclass(frozen=True, slots=True)
class _CompatibilityProfile:
    profile_id: str
    document_server_version: str
    connector_version: str
    validation_formats: frozenset[FileFormat]
    rules: tuple[_KnownDriftRule, ...]
    evidence_scope: str = _EVIDENCE_SCOPE


@dataclass(frozen=True, slots=True)
class _ClassifiedFinding:
    finding: ValidationError
    rule_id: str | None


@dataclass(frozen=True, slots=True)
class EuroOfficeCompatibilityReport:
    """Compatibility view that retains the raw strict outcome alongside it."""

    status: EuroOfficeCompatibilityStatus
    profile_id: str
    observed_document_server_version: str
    observed_connector_version: str
    expected_document_server_version: str
    expected_connector_version: str
    evidence_scope: str
    strict_validation: bool
    security_validation: bool
    complete_scan: bool
    strict_valid: bool
    strict_error_count: int
    strict_warning_count: int
    accepted_findings: tuple[_ClassifiedFinding, ...]
    unexpected_findings: tuple[_ClassifiedFinding, ...]
    rule_counts: tuple[tuple[str, int], ...]
    rule_limits: tuple[tuple[str, int], ...]
    exceeded_rules: tuple[str, ...]
    mismatch_reasons: tuple[str, ...]
    semantic_preservation: str = _SEMANTIC_PRESERVATION

    @property
    def compatible(self) -> bool:
        """Whether this exact profile accepts the static validation outcome."""
        return self.status in {
            EuroOfficeCompatibilityStatus.STRICT_CLEAN,
            EuroOfficeCompatibilityStatus.ACCEPTED_KNOWN_DRIFT,
        }

    @property
    def accepted_finding_count(self) -> int:
        return len(self.accepted_findings)

    @property
    def unexpected_finding_count(self) -> int:
        return len(self.unexpected_findings)

    def to_dict(self) -> dict[str, Any]:
        """Return a compact report suitable for Nextcloud quality/compliance JSON."""
        return {
            "status": self.status.value,
            "compatible": self.compatible,
            "profile_id": self.profile_id,
            "environment": {
                "observed_document_server_version": self.observed_document_server_version,
                "observed_connector_version": self.observed_connector_version,
                "expected_document_server_version": self.expected_document_server_version,
                "expected_connector_version": self.expected_connector_version,
            },
            "evidence_scope": self.evidence_scope,
            "semantic_preservation": self.semantic_preservation,
            "validation_contract": {
                "strict": self.strict_validation,
                "security_validation": self.security_validation,
                "complete_scan": self.complete_scan,
            },
            "raw_strict": {
                "valid": self.strict_valid,
                "error_count": self.strict_error_count,
                "warning_count": self.strict_warning_count,
            },
            "accepted_finding_count": self.accepted_finding_count,
            "unexpected_finding_count": self.unexpected_finding_count,
            "rule_counts": dict(self.rule_counts),
            "rule_limits": dict(self.rule_limits),
            "exceeded_rules": list(self.exceeded_rules),
            "mismatch_reasons": list(self.mismatch_reasons),
            "unexpected_finding_groups": _summarize_findings(self.unexpected_findings),
        }


def _rule(
    rule_id: str,
    *,
    extensions: frozenset[str],
    severity: ValidationSeverity,
    error_type: ValidationErrorType,
    source_class: SourceClass,
    error_id: str,
    part_uri: str,
    node: str | None,
    description: str,
    max_occurrences: int,
    formats: frozenset[FileFormat] = _OOXML_FORMATS,
    paths: frozenset[str] | None = None,
    path_pattern: re.Pattern[str] | None = None,
    requires_no_bin_payload: bool = False,
) -> _KnownDriftRule:
    return _KnownDriftRule(
        rule_id=rule_id,
        extensions=extensions,
        formats=formats,
        severity=severity,
        error_type=error_type,
        source_class=source_class,
        error_id=error_id,
        part_uri=part_uri,
        node=node,
        description=description,
        max_occurrences=max_occurrences,
        paths=paths,
        path_pattern=path_pattern,
        requires_no_bin_payload=requires_no_bin_payload,
    )


_ALL_CORE_EXTENSIONS = frozenset({".docx", ".pptx", ".xlsx"})
_DOCX = frozenset({".docx"})
_XLSX = frozenset({".xlsx"})
_DOCX_XLSX = frozenset({".docx", ".xlsx"})

_UNEXPECTED_ELEMENT_ID = "Sch_UnexpectedElementContentExpectingComplex"
_UNEXPECTED_ELEMENT_DESCRIPTION = "Unexpected element '{node}' found"

_RULES = (
    _rule(
        "content-types.ole-bin-default-without-payload",
        extensions=_ALL_CORE_EXTENSIONS,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SEMANTIC,
        source_class=SourceClass.SDK_PROXY,
        error_id="Sec_ActiveContentType",
        part_uri="/[Content_Types].xml",
        node="Default",
        description=(
            "Active content content type "
            "'application/vnd.openxmlformats-officedocument.oleObject' declared for extension "
            "'.bin'"
        ),
        max_occurrences=1,
        path_pattern=re.compile(r"/Types\[\d+\]/Default\[\d+\]"),
        requires_no_bin_payload=True,
    ),
    _rule(
        "extended-properties.empty-heading-vectors",
        extensions=_DOCX_XLSX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id="Sch_IncompleteContentExpectingComplex",
        part_uri="/docProps/app.xml",
        node=None,
        description="Required choice element is missing",
        max_occurrences=2,
        paths=frozenset(
            {
                "/Properties[1]/HeadingPairs[1]/vector[1]",
                "/Properties[1]/TitlesOfParts[1]/vector[1]",
            }
        ),
    ),
    _rule(
        "word.styles.table-style-paragraph-properties",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/word/styles.xml",
        node="pPr",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="pPr"),
        max_occurrences=1411,
        path_pattern=re.compile(r"/styles\[\d+\]/style\[\d+\]/tblStylePr\[\d+\]"),
    ),
    _rule(
        "word.styles.table-cell-borders",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/word/styles.xml",
        node="tcBorders",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="tcBorders"),
        max_occurrences=1296,
        path_pattern=re.compile(r"/styles\[\d+\]/style\[\d+\]/(?:tblStylePr\[\d+\]/)?tcPr\[\d+\]"),
    ),
    _rule(
        "word.styles.table-borders",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/word/styles.xml",
        node="tblBorders",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="tblBorders"),
        max_occurrences=187,
        path_pattern=re.compile(r"/styles\[\d+\]/style\[\d+\]/tblPr\[\d+\]"),
    ),
    _rule(
        "word.styles.table-cell-margin-top",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/word/styles.xml",
        node="top",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="top"),
        max_occurrences=187,
        path_pattern=re.compile(r"/styles\[\d+\]/style\[\d+\]/tblPr\[\d+\]/tblCellMar\[\d+\]"),
    ),
    _rule(
        "word.numbering.level-text-order",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/word/numbering.xml",
        node="lvlText",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="lvlText"),
        max_occurrences=18,
        path_pattern=re.compile(r"/numbering\[\d+\]/abstractNum\[\d+\]/lvl\[\d+\]"),
    ),
    _rule(
        "word.document.table-cell-width-order",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/word/document.xml",
        node="tcW",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="tcW"),
        max_occurrences=4,
        path_pattern=re.compile(
            r"/document\[\d+\]/body\[\d+\]/tbl\[\d+\]/tr\[\d+\]/"
            r"tc\[\d+\]/tcPr\[\d+\]"
        ),
    ),
    _rule(
        "word.document.table-cell-width-app-compat",
        extensions=_DOCX,
        severity=ValidationSeverity.WARNING,
        error_type=ValidationErrorType.SEMANTIC,
        source_class=SourceClass.WORD_APP_COMPAT,
        error_id="Sem_SemanticError",
        part_uri="/word/document.xml",
        node="tcW",
        description=(
            "tcPr child 'tcW' appears after 'tcBorders' but ECMA-376 §17.4.70 places it "
            "earlier — Word may flag this file as unreadable content"
        ),
        max_occurrences=4,
        paths=frozenset({"/document[1]"}),
    ),
    _rule(
        "word.settings.endnote-position-doc-end",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id="Sch_AttributeValueDataTypeDetailed",
        part_uri="/word/settings.xml",
        node="val",
        description=(
            "Invalid value for attribute 'val': Value 'docEnd' is not in allowed values: "
            "['beneathText', 'pageBottom', 'sectEnd']"
        ),
        max_occurrences=1,
        paths=frozenset({"/settings[1]/endnotePr[1]/pos[1]"}),
    ),
    _rule(
        "excel.chart.data-label-category-name-order",
        extensions=_XLSX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id=_UNEXPECTED_ELEMENT_ID,
        part_uri="/xl/charts/chart1.xml",
        node="showCatName",
        description=_UNEXPECTED_ELEMENT_DESCRIPTION.format(node="showCatName"),
        max_occurrences=1,
        paths=frozenset({"/chartSpace[1]/chart[1]/plotArea[1]/barChart[1]/dLbls[1]"}),
    ),
)

_OFFICE_2007_TABLE_LOOK_RULES = tuple(
    _rule(
        f"word.document.table-look-{attribute}",
        extensions=_DOCX,
        severity=ValidationSeverity.ERROR,
        error_type=ValidationErrorType.SCHEMA,
        source_class=SourceClass.SDK_PROXY,
        error_id="Sch_UndeclaredAttribute",
        part_uri="/word/document.xml",
        node=attribute,
        description=f"The '{attribute}' attribute is not declared.",
        max_occurrences=1,
        formats=frozenset({FileFormat.OFFICE_2007}),
        paths=frozenset({"/document[1]/body[1]/tbl[1]/tblPr[1]/tblLook[1]"}),
    )
    for attribute in (
        "firstRow",
        "lastRow",
        "firstColumn",
        "lastColumn",
        "noHBand",
        "noVBand",
    )
)

_PROFILE = _CompatibilityProfile(
    profile_id=TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
    document_server_version="9.3.1.37",
    connector_version="11.0.1",
    validation_formats=_OOXML_FORMATS,
    rules=_RULES + _OFFICE_2007_TABLE_LOOK_RULES,
)

_PROFILES = {_PROFILE.profile_id: _PROFILE}


def supported_eurooffice_compatibility_profiles() -> tuple[str, ...]:
    """Return the compatibility profile identifiers shipped by this release."""
    return tuple(sorted(_PROFILES))


def _package_bin_payload_state(file_path: str) -> bool | None:
    """Return True/False for a .bin member, or None when it cannot be verified."""
    path = Path(file_path)
    if not path.is_file():
        return None
    try:
        with ZipFile(path) as package:
            return any(
                not name.endswith("/") and Path(name).suffix.lower() == ".bin"
                for name in package.namelist()
            )
    except (BadZipFile, OSError):
        return None


def _summarize_findings(findings: tuple[_ClassifiedFinding, ...]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, ...], dict[str, Any]] = {}
    for classified in findings:
        finding = classified.finding
        key = (
            classified.rule_id or "",
            normalize_error_id(finding),
            finding.error_type.value,
            finding.severity.value,
            finding.source_class.value,
            finding.part_uri,
            finding.node or "",
            finding.description,
        )
        if key not in groups:
            groups[key] = {
                "rule_id": classified.rule_id,
                "id": normalize_error_id(finding),
                "error_type": finding.error_type.value,
                "severity": finding.severity.value,
                "source_class": finding.source_class.value,
                "part_uri": finding.part_uri,
                "node": finding.node,
                "description": finding.description,
                "example_path": finding.path,
                "count": 0,
            }
        groups[key]["count"] += 1
    return list(groups.values())


def _environment_mismatches(
    result: ValidationResult,
    profile: _CompatibilityProfile,
    *,
    document_server_version: str,
    connector_version: str,
    strict_validation: bool,
    security_validation: bool,
    complete_scan: bool,
) -> tuple[str, ...]:
    reasons: list[str] = []
    if document_server_version != profile.document_server_version:
        reasons.append(
            "document server version mismatch: "
            f"expected {profile.document_server_version}, observed {document_server_version}"
        )
    if connector_version != profile.connector_version:
        reasons.append(
            "connector version mismatch: "
            f"expected {profile.connector_version}, observed {connector_version}"
        )
    if result.file_format not in profile.validation_formats:
        supported = ", ".join(
            sorted(file_format.value for file_format in profile.validation_formats)
        )
        reasons.append(
            f"validation format {result.file_format.value} was not calibrated; "
            f"supported: {supported}"
        )
    extension = Path(result.file_path).suffix.lower()
    supported_extensions = frozenset(
        extension for rule in profile.rules for extension in rule.extensions
    )
    if extension not in supported_extensions:
        supported = ", ".join(sorted(supported_extensions))
        reasons.append(
            f"file extension {extension or '<none>'} was not calibrated; supported: {supported}"
        )
    if not strict_validation:
        reasons.append("validation was not run in strict mode")
    if not security_validation:
        reasons.append("OOXML security validation was not enabled")
    if not complete_scan:
        reasons.append("validation finding collection was not confirmed complete")
    return tuple(reasons)


def classify_eurooffice_compatibility(
    result: ValidationResult,
    *,
    document_server_version: str,
    connector_version: str,
    strict_validation: bool = False,
    security_validation: bool = False,
    complete_scan: bool = False,
    profile_id: str = TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
) -> EuroOfficeCompatibilityReport:
    """Classify raw findings against one exact EuroOffice editor profile.

    Unknown profiles are caller configuration errors.  Known profiles with a
    mismatched runtime version or validation contract return
    ``unverified-environment`` and accept no findings. Callers must explicitly
    attest that the supplied result came from strict, security-enabled,
    unlimited finding collection; the CLI establishes that contract
    automatically in profile mode.
    """
    try:
        profile = _PROFILES[profile_id]
    except KeyError as exc:
        supported = ", ".join(supported_eurooffice_compatibility_profiles())
        raise ValueError(
            f"Unknown EuroOffice compatibility profile {profile_id!r}; {supported}"
        ) from exc

    mismatch_reasons = _environment_mismatches(
        result,
        profile,
        document_server_version=document_server_version,
        connector_version=connector_version,
        strict_validation=strict_validation,
        security_validation=security_validation,
        complete_scan=complete_scan,
    )
    if mismatch_reasons:
        unexpected = tuple(_ClassifiedFinding(finding, None) for finding in result.errors)
        return EuroOfficeCompatibilityReport(
            status=EuroOfficeCompatibilityStatus.UNVERIFIED_ENVIRONMENT,
            profile_id=profile.profile_id,
            observed_document_server_version=document_server_version,
            observed_connector_version=connector_version,
            expected_document_server_version=profile.document_server_version,
            expected_connector_version=profile.connector_version,
            evidence_scope=profile.evidence_scope,
            strict_validation=strict_validation,
            security_validation=security_validation,
            complete_scan=complete_scan,
            strict_valid=result.is_valid,
            strict_error_count=result.error_count,
            strict_warning_count=result.warning_count,
            accepted_findings=(),
            unexpected_findings=unexpected,
            rule_counts=(),
            rule_limits=(),
            exceeded_rules=(),
            mismatch_reasons=mismatch_reasons,
        )

    extension = Path(result.file_path).suffix.lower()
    matched: dict[str, list[ValidationError]] = {rule.rule_id: [] for rule in profile.rules}
    unmatched: list[_ClassifiedFinding] = []
    rule_by_id = {rule.rule_id: rule for rule in profile.rules}

    for finding in result.errors:
        rule = next(
            (
                candidate
                for candidate in profile.rules
                if candidate.matches(
                    finding,
                    extension=extension,
                    file_format=result.file_format,
                )
            ),
            None,
        )
        if rule is None:
            unmatched.append(_ClassifiedFinding(finding, None))
        else:
            matched[rule.rule_id].append(finding)

    accepted: list[_ClassifiedFinding] = []
    rejected: list[_ClassifiedFinding] = unmatched
    exceeded_rules: list[str] = []
    rejected_rule_reasons: list[str] = []
    rule_counts = Counter(
        {rule_id: len(findings) for rule_id, findings in matched.items() if findings}
    )
    bin_payload_state: bool | None = None

    for rule_id, findings in matched.items():
        if not findings:
            continue
        rule = rule_by_id[rule_id]
        reject_reason: str | None = None
        if len(findings) > rule.max_occurrences:
            exceeded_rules.append(rule_id)
            reject_reason = (
                f"{rule_id}: observed {len(findings)}, profile maximum {rule.max_occurrences}"
            )
        elif rule.requires_no_bin_payload:
            if bin_payload_state is None:
                bin_payload_state = _package_bin_payload_state(result.file_path)
            if bin_payload_state is True:
                reject_reason = f"{rule_id}: package contains a .bin payload"
            elif bin_payload_state is None:
                reject_reason = f"{rule_id}: package payload could not be inspected"

        classified = [_ClassifiedFinding(finding, rule_id) for finding in findings]
        if reject_reason is None:
            accepted.extend(classified)
        else:
            rejected.extend(classified)
            rejected_rule_reasons.append(reject_reason)

    if rejected:
        status = EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    elif accepted:
        status = EuroOfficeCompatibilityStatus.ACCEPTED_KNOWN_DRIFT
    else:
        status = EuroOfficeCompatibilityStatus.STRICT_CLEAN

    return EuroOfficeCompatibilityReport(
        status=status,
        profile_id=profile.profile_id,
        observed_document_server_version=document_server_version,
        observed_connector_version=connector_version,
        expected_document_server_version=profile.document_server_version,
        expected_connector_version=profile.connector_version,
        evidence_scope=profile.evidence_scope,
        strict_validation=strict_validation,
        security_validation=security_validation,
        complete_scan=complete_scan,
        strict_valid=result.is_valid,
        strict_error_count=result.error_count,
        strict_warning_count=result.warning_count,
        accepted_findings=tuple(accepted),
        unexpected_findings=tuple(rejected),
        rule_counts=tuple(sorted(rule_counts.items())),
        rule_limits=tuple(
            sorted((rule_id, rule_by_id[rule_id].max_occurrences) for rule_id in rule_counts)
        ),
        exceeded_rules=tuple(sorted(exceeded_rules)),
        mismatch_reasons=tuple(rejected_rule_reasons),
    )
