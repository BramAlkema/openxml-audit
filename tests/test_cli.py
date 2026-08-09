"""CLI routing tests for validator auto-detection."""

from __future__ import annotations

import json
from pathlib import Path

from click.testing import CliRunner

from openxml_audit import ValidationError, ValidationResult
from openxml_audit import cli as cli_module
from openxml_audit.errors import FileFormat, ValidationErrorType, ValidationSeverity
from openxml_audit.eurooffice.compatibility import TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID


def test_detect_validator_for_odf_extensions() -> None:
    assert cli_module._detect_validator_for_path(Path("sample.odt")) == "odf"
    assert cli_module._detect_validator_for_path(Path("sample.ods")) == "odf"
    assert cli_module._detect_validator_for_path(Path("sample.odp")) == "odf"


def test_cli_auto_routes_odf_to_odf_validator(monkeypatch, minimal_odt: Path) -> None:
    calls: list[tuple[str, str, FileFormat, int, bool, dict[str, object]]] = []

    class DummyOdfValidator:
        def __init__(
            self,
            file_format: FileFormat,
            max_errors: int,
            strict: bool,
            **kwargs: object,
        ) -> None:
            self.file_format = file_format
            self.max_errors = max_errors
            self.strict = strict
            self.kwargs = kwargs

        def validate(self, path: Path) -> ValidationResult:
            calls.append(
                (
                    "odf",
                    str(path),
                    self.file_format,
                    self.max_errors,
                    self.strict,
                    self.kwargs,
                )
            )
            return ValidationResult(
                is_valid=True,
                errors=[],
                file_path=str(path),
                file_format=self.file_format,
            )

    class FailOpenXmlValidator:
        def __init__(self, *args, **kwargs) -> None:  # type: ignore[no-untyped-def]
            raise AssertionError("OpenXmlValidator should not be constructed for .odt")

    monkeypatch.setattr(cli_module, "OdfValidator", DummyOdfValidator)
    monkeypatch.setattr(cli_module, "OpenXmlValidator", FailOpenXmlValidator)

    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [str(minimal_odt), "--validator", "auto", "--output", "json"],
    )

    assert result.exit_code == 0, result.output
    assert calls and calls[0][0] == "odf"
    assert calls[0][2] == FileFormat.ODF_1_3
    assert calls[0][3] == 100
    assert calls[0][5]["semantic_validation"] is True


def test_cli_odf_level_foundation_disables_semantic_schema(
    monkeypatch,
    minimal_odt: Path,
) -> None:
    kwargs_capture: list[dict[str, object]] = []

    class DummyOdfValidator:
        def __init__(
            self,
            file_format: FileFormat,
            max_errors: int,
            strict: bool,
            **kwargs: object,
        ) -> None:
            kwargs_capture.append(kwargs)
            self.file_format = file_format

        def validate(self, path: Path) -> ValidationResult:
            return ValidationResult(
                is_valid=True,
                errors=[],
                file_path=str(path),
                file_format=self.file_format,
            )

    monkeypatch.setattr(cli_module, "OdfValidator", DummyOdfValidator)

    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [
            str(minimal_odt),
            "--validator",
            "odf",
            "--odf-level",
            "foundation",
            "--output",
            "json",
        ],
    )

    assert result.exit_code == 0, result.output
    assert kwargs_capture
    kwargs = kwargs_capture[0]
    assert kwargs["schema_validation"] is False
    assert kwargs["semantic_validation"] is False
    assert kwargs["security_validation"] is False
    assert kwargs["relaxng_validation"] is False


def test_cli_odf_schema_core_uses_bundled_routes_by_default(
    monkeypatch,
    minimal_odt: Path,
) -> None:
    kwargs_capture: list[dict[str, object]] = []

    class DummyOdfValidator:
        def __init__(
            self,
            file_format: FileFormat,
            max_errors: int,
            strict: bool,
            **kwargs: object,
        ) -> None:
            kwargs_capture.append(kwargs)
            self.file_format = file_format

        def validate(self, path: Path) -> ValidationResult:
            return ValidationResult(
                is_valid=True,
                errors=[],
                file_path=str(path),
                file_format=self.file_format,
            )

    monkeypatch.setattr(cli_module, "OdfValidator", DummyOdfValidator)

    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [
            str(minimal_odt),
            "--validator",
            "odf",
            "--odf-level",
            "schema-core",
            "--output",
            "json",
        ],
    )

    assert result.exit_code == 0, result.output
    assert kwargs_capture
    kwargs = kwargs_capture[0]
    assert kwargs["schema_validation"] is True
    assert kwargs["semantic_validation"] is False
    assert kwargs["security_validation"] is False
    assert kwargs["relaxng_validation"] is True
    assert kwargs["schema_routes"] is None
    assert kwargs["require_schema_routes"] is False


def test_cli_auto_routes_ooxml_to_ooxml_validator(monkeypatch, minimal_pptx: Path) -> None:
    calls: list[tuple[str, str, FileFormat, bool, bool]] = []

    class DummyOpenXmlValidator:
        def __init__(
            self,
            file_format: FileFormat,
            max_errors: int,
            strict: bool,
            security_validation: bool,
        ) -> None:
            self.file_format = file_format
            self.strict = strict
            self.security_validation = security_validation

        def validate(self, path: Path) -> ValidationResult:
            calls.append(
                (
                    "ooxml",
                    str(path),
                    self.file_format,
                    self.strict,
                    self.security_validation,
                )
            )
            return ValidationResult(
                is_valid=True,
                errors=[],
                file_path=str(path),
                file_format=self.file_format,
            )

    class FailOdfValidator:
        def __init__(self, *args, **kwargs) -> None:  # type: ignore[no-untyped-def]
            raise AssertionError("OdfValidator should not be constructed for .pptx")

    monkeypatch.setattr(cli_module, "OpenXmlValidator", DummyOpenXmlValidator)
    monkeypatch.setattr(cli_module, "OdfValidator", FailOdfValidator)

    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [str(minimal_pptx), "--validator", "auto", "--output", "json"],
    )

    assert result.exit_code == 0, result.output
    assert calls and calls[0][0] == "ooxml"
    assert calls[0][2] == FileFormat.OFFICE_2019
    assert calls[0][4] is False


def test_cli_can_enable_ooxml_security(monkeypatch, minimal_pptx: Path) -> None:
    calls: list[tuple[FileFormat, bool]] = []

    class DummyOpenXmlValidator:
        def __init__(
            self,
            file_format: FileFormat,
            max_errors: int,
            strict: bool,
            security_validation: bool,
        ) -> None:
            del max_errors, strict
            calls.append((file_format, security_validation))
            self.file_format = file_format

        def validate(self, path: Path) -> ValidationResult:
            return ValidationResult(
                is_valid=True,
                errors=[],
                file_path=str(path),
                file_format=self.file_format,
            )

    monkeypatch.setattr(cli_module, "OpenXmlValidator", DummyOpenXmlValidator)

    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [
            str(minimal_pptx),
            "--validator",
            "ooxml",
            "--ooxml-security",
            "--output",
            "json",
        ],
    )

    assert result.exit_code == 0, result.output
    assert calls == [(FileFormat.OFFICE_2019, True)]


def test_cli_eurooffice_profile_forces_full_secure_scan_and_adds_json_report(
    monkeypatch,
    minimal_pptx: Path,
) -> None:
    calls: list[tuple[FileFormat, int, bool]] = []

    class DummyOpenXmlValidator:
        def __init__(
            self,
            file_format: FileFormat,
            max_errors: int,
            strict: bool,
            security_validation: bool,
        ) -> None:
            del strict
            calls.append((file_format, max_errors, security_validation))
            self.file_format = file_format

        def validate(self, path: Path) -> ValidationResult:
            return ValidationResult(
                is_valid=False,
                errors=[
                    ValidationError(
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            "Active content content type "
                            "'application/vnd.openxmlformats-officedocument.oleObject' "
                            "declared for extension '.bin'"
                        ),
                        part_uri="/[Content_Types].xml",
                        path="/Types[1]/Default[6]",
                        node="Default",
                        id="Sec_ActiveContentType",
                    )
                ],
                file_path=str(path),
                file_format=self.file_format,
            )

    monkeypatch.setattr(cli_module, "OpenXmlValidator", DummyOpenXmlValidator)

    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [
            str(minimal_pptx),
            "--validator",
            "ooxml",
            "--output",
            "json",
            "--max-errors",
            "1",
            "--no-ooxml-security",
            "--eurooffice-profile",
            TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
            "--eurooffice-document-server-version",
            "9.3.1.37",
            "--eurooffice-connector-version",
            "11.0.1",
        ],
    )

    # Compatibility mode is additive: accepted drift is in JSON, while the
    # existing strict exit status and raw invalid result remain unchanged.
    assert result.exit_code == 1, result.output
    assert calls == [(FileFormat.MICROSOFT_365, 0, True)]
    payload = json.loads(result.output)
    assert payload[0]["valid"] is False
    compatibility = payload[0]["eurooffice_compatibility"]
    assert compatibility["status"] == "accepted-known-drift"
    assert compatibility["compatible"] is True
    assert compatibility["raw_strict"]["error_count"] == 1
    assert compatibility["semantic_preservation"] == "not-assessed"


def test_cli_eurooffice_profile_requires_observed_versions(minimal_pptx: Path) -> None:
    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [
            str(minimal_pptx),
            "--eurooffice-profile",
            TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
        ],
    )

    assert result.exit_code == 1
    assert "requires --eurooffice-document-server-version" in result.output


def test_cli_eurooffice_profile_rejects_permissive_policy(minimal_pptx: Path) -> None:
    runner = CliRunner()
    result = runner.invoke(
        cli_module.main,
        [
            str(minimal_pptx),
            "--policy",
            "permissive",
            "--eurooffice-profile",
            TOKENOOXML_EUROOFFICE_EDITOR_PROFILE_ID,
            "--eurooffice-document-server-version",
            "9.3.1.37",
            "--eurooffice-connector-version",
            "11.0.1",
        ],
    )

    assert result.exit_code == 1
    assert "require --policy=strict" in result.output


def test_output_text_includes_warnings() -> None:
    result = ValidationResult(
        is_valid=True,
        errors=[
            ValidationError(
                error_type=ValidationErrorType.SEMANTIC,
                severity=ValidationSeverity.WARNING,
                description="External relationship target: https://example.com",
                part_uri="/ppt/slides/_rels/slide1.xml.rels",
            )
        ],
        file_path="sample.pptx",
        file_format=FileFormat.OFFICE_2019,
    )

    with cli_module.console.capture() as capture:
        cli_module._output_text([result], quiet=False)

    output = capture.get()
    assert "External relationship target: https://example.com" in output
    assert "Warnings: 1" in output
