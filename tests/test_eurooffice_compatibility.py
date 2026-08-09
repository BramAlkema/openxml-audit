"""Tests for the version-bound EuroOffice editor compatibility profile."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from openxml_audit.errors import (
    FileFormat,
    SourceClass,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.eurooffice.compatibility import (
    TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
    EuroOfficeCompatibilityStatus,
    classify_eurooffice_compatibility,
    supported_eurooffice_compatibility_profiles,
)

DOCUMENT_SERVER_VERSION = "9.3.1.37"
CONNECTOR_VERSION = "11.0.1"


def _package(tmp_path: Path, extension: str = ".docx", *, with_bin: bool = False) -> Path:
    path = tmp_path / f"callback{extension}"
    with ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types/>")
        if with_bin:
            package.writestr("word/embeddings/oleObject1.bin", b"payload")
    return path


def _finding(
    *,
    error_type: ValidationErrorType,
    description: str,
    part_uri: str,
    node: str | None,
    path: str = "",
    severity: ValidationSeverity = ValidationSeverity.ERROR,
    source_class: SourceClass = SourceClass.SDK_PROXY,
    error_id: str = "",
) -> ValidationError:
    return ValidationError(
        error_type=error_type,
        description=description,
        part_uri=part_uri,
        path=path,
        node=node,
        severity=severity,
        id=error_id,
        source_class=source_class,
    )


def _result(
    file_path: Path,
    findings: list[ValidationError],
    *,
    file_format: FileFormat = FileFormat.MICROSOFT_365,
) -> ValidationResult:
    return ValidationResult(
        is_valid=not any(finding.severity is ValidationSeverity.ERROR for finding in findings),
        errors=findings,
        file_path=str(file_path),
        file_format=file_format,
    )


def _classify(result: ValidationResult):
    return classify_eurooffice_compatibility(
        result,
        document_server_version=DOCUMENT_SERVER_VERSION,
        connector_version=CONNECTOR_VERSION,
        strict_validation=True,
        security_validation=True,
        complete_scan=True,
    )


def _phantom_ole_finding() -> ValidationError:
    return _finding(
        error_type=ValidationErrorType.SEMANTIC,
        description=(
            "Active content content type "
            "'application/vnd.openxmlformats-officedocument.oleObject' declared for extension "
            "'.bin'"
        ),
        part_uri="/[Content_Types].xml",
        path="/Types[1]/Default[6]",
        node="Default",
        error_id="Sec_ActiveContentType",
    )


def _word_ppr_finding() -> ValidationError:
    return _finding(
        error_type=ValidationErrorType.SCHEMA,
        description="Unexpected element 'pPr' found",
        part_uri="/word/styles.xml",
        path="/styles[1]/style[2]/tblStylePr[5]",
        node="pPr",
    )


def _word_order_warning() -> ValidationError:
    return _finding(
        error_type=ValidationErrorType.SEMANTIC,
        description=(
            "tcPr child 'tcW' appears after 'tcBorders' but ECMA-376 §17.4.70 places it "
            "earlier — Word may flag this file as unreadable content"
        ),
        part_uri="/word/document.xml",
        path="/document[1]",
        node="tcW",
        severity=ValidationSeverity.WARNING,
        source_class=SourceClass.WORD_APP_COMPAT,
    )


def test_supported_profile_is_explicit_and_version_bound():
    assert supported_eurooffice_compatibility_profiles() == (
        TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
    )


def test_known_drift_is_accepted_without_changing_raw_strict_result(tmp_path: Path):
    path = _package(tmp_path)
    findings = [_phantom_ole_finding(), _word_ppr_finding(), _word_order_warning()]
    result = _result(path, findings)

    report = _classify(result)

    assert report.status is EuroOfficeCompatibilityStatus.ACCEPTED_KNOWN_DRIFT
    assert report.compatible is True
    assert report.strict_valid is False
    assert report.strict_error_count == 2
    assert report.strict_warning_count == 1
    assert report.accepted_finding_count == 3
    assert report.unexpected_finding_count == 0
    assert report.semantic_preservation == "not-assessed"
    assert result.errors is findings
    assert result.is_valid is False

    payload = report.to_dict()
    assert payload["raw_strict"] == {
        "valid": False,
        "error_count": 2,
        "warning_count": 1,
    }
    assert payload["validation_contract"] == {
        "strict": True,
        "security_validation": True,
        "complete_scan": True,
    }
    assert payload["semantic_preservation"] == "not-assessed"
    assert payload["unexpected_finding_groups"] == []
    assert payload["rule_limits"] == {
        "content-types.ole-bin-default-without-payload": 1,
        "word.document.table-cell-width-app-compat": 4,
        "word.styles.table-style-paragraph-properties": 1411,
    }


def test_strict_clean_result_needs_no_waiver(tmp_path: Path):
    report = _classify(_result(_package(tmp_path, ".pptx"), []))

    assert report.status is EuroOfficeCompatibilityStatus.STRICT_CLEAN
    assert report.compatible is True
    assert report.accepted_finding_count == 0


def test_direct_api_requires_explicit_complete_secure_strict_scan_contract(tmp_path: Path):
    report = classify_eurooffice_compatibility(
        _result(_package(tmp_path), []),
        document_server_version=DOCUMENT_SERVER_VERSION,
        connector_version=CONNECTOR_VERSION,
    )

    assert report.status is EuroOfficeCompatibilityStatus.UNVERIFIED_ENVIRONMENT
    assert report.compatible is False
    assert report.mismatch_reasons == (
        "validation was not run in strict mode",
        "OOXML security validation was not enabled",
        "validation finding collection was not confirmed complete",
    )


def test_uncalibrated_file_extension_is_unverified(tmp_path: Path):
    report = _classify(_result(_package(tmp_path, ".docm"), []))

    assert report.status is EuroOfficeCompatibilityStatus.UNVERIFIED_ENVIRONMENT
    assert "file extension .docm was not calibrated" in report.mismatch_reasons[0]


def test_rule_count_at_observed_maximum_is_accepted(tmp_path: Path):
    finding = _word_ppr_finding()
    report = _classify(_result(_package(tmp_path), [finding] * 1411))

    assert report.status is EuroOfficeCompatibilityStatus.ACCEPTED_KNOWN_DRIFT
    assert report.accepted_finding_count == 1411
    assert report.exceeded_rules == ()


def test_rule_count_above_observed_maximum_is_unexpected(tmp_path: Path):
    finding = _word_ppr_finding()
    result = _result(_package(tmp_path), [finding] * 1412)

    report = _classify(result)

    assert report.status is EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    assert report.compatible is False
    assert report.accepted_finding_count == 0
    assert report.unexpected_finding_count == 1412
    assert report.exceeded_rules == ("word.styles.table-style-paragraph-properties",)
    assert report.to_dict()["unexpected_finding_groups"][0]["count"] == 1412


def test_ole_declaration_is_not_waived_when_package_contains_bin_payload(tmp_path: Path):
    result = _result(_package(tmp_path, with_bin=True), [_phantom_ole_finding()])

    report = _classify(result)

    assert report.status is EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    assert report.accepted_finding_count == 0
    assert report.unexpected_finding_count == 1
    assert report.mismatch_reasons == (
        "content-types.ole-bin-default-without-payload: package contains a .bin payload",
    )


def test_ole_declaration_is_not_waived_when_package_cannot_be_inspected(tmp_path: Path):
    missing_path = tmp_path / "missing.docx"
    report = _classify(_result(missing_path, [_phantom_ole_finding()]))

    assert report.status is EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    assert report.mismatch_reasons == (
        "content-types.ole-bin-default-without-payload: package payload could not be inspected",
    )


def test_unknown_finding_family_remains_unexpected(tmp_path: Path):
    unknown = _finding(
        error_type=ValidationErrorType.SEMANTIC,
        description="Content control was removed",
        part_uri="/word/document.xml",
        node="sdt",
        error_id="Sem_SemanticError",
    )

    report = _classify(_result(_package(tmp_path), [unknown]))

    assert report.status is EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    assert report.accepted_finding_count == 0
    assert report.unexpected_findings[0].rule_id is None


def test_known_family_at_unobserved_path_remains_unexpected(tmp_path: Path):
    near_match = _word_ppr_finding()
    near_match.path = "/styles[1]/style[2]/pPr[1]"

    report = _classify(_result(_package(tmp_path), [near_match]))

    assert report.status is EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    assert report.accepted_finding_count == 0
    assert report.unexpected_findings[0].rule_id is None


@pytest.mark.parametrize(
    ("document_server_version", "connector_version"),
    [
        ("9.3.2.0", CONNECTOR_VERSION),
        (DOCUMENT_SERVER_VERSION, "11.1.0"),
    ],
)
def test_unknown_runtime_version_accepts_no_findings(
    tmp_path: Path,
    document_server_version: str,
    connector_version: str,
):
    result = _result(_package(tmp_path), [_word_ppr_finding()])

    report = classify_eurooffice_compatibility(
        result,
        document_server_version=document_server_version,
        connector_version=connector_version,
        strict_validation=True,
        security_validation=True,
        complete_scan=True,
    )

    assert report.status is EuroOfficeCompatibilityStatus.UNVERIFIED_ENVIRONMENT
    assert report.compatible is False
    assert report.accepted_finding_count == 0
    assert report.unexpected_finding_count == 1
    assert report.mismatch_reasons


def test_uncalibrated_validation_format_is_unverified(tmp_path: Path):
    result = _result(
        _package(tmp_path),
        [_word_ppr_finding()],
        file_format=FileFormat.OFFICE_2019,
    )

    report = _classify(result)

    assert report.status is EuroOfficeCompatibilityStatus.UNVERIFIED_ENVIRONMENT
    assert "office2019 was not calibrated" in report.mismatch_reasons[0]


def test_office_2007_table_look_attribute_is_not_accepted_for_microsoft_365(
    tmp_path: Path,
):
    finding = _finding(
        error_type=ValidationErrorType.SCHEMA,
        description="The 'firstRow' attribute is not declared.",
        part_uri="/word/document.xml",
        path="/document[1]/body[1]/tbl[1]/tblPr[1]/tblLook[1]",
        node="firstRow",
    )
    microsoft_365 = _classify(_result(_package(tmp_path), [finding]))

    office_2007_result = _result(
        _package(tmp_path),
        [finding],
        file_format=FileFormat.OFFICE_2007,
    )
    office_2007 = _classify(office_2007_result)

    assert microsoft_365.status is EuroOfficeCompatibilityStatus.UNEXPECTED_DRIFT
    assert office_2007.status is EuroOfficeCompatibilityStatus.ACCEPTED_KNOWN_DRIFT


def test_unknown_profile_is_a_caller_error(tmp_path: Path):
    with pytest.raises(ValueError, match="Unknown EuroOffice compatibility profile"):
        classify_eurooffice_compatibility(
            _result(_package(tmp_path), []),
            profile_id="unknown",
            document_server_version=DOCUMENT_SERVER_VERSION,
            connector_version=CONNECTOR_VERSION,
            strict_validation=True,
            security_validation=True,
            complete_scan=True,
        )
