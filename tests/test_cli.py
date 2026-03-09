"""CLI routing tests for validator auto-detection."""

from __future__ import annotations

from pathlib import Path

from click.testing import CliRunner

from openxml_audit import ValidationResult
from openxml_audit import cli as cli_module
from openxml_audit.errors import FileFormat


def test_detect_validator_for_odf_extensions() -> None:
    assert cli_module._detect_validator_for_path(Path("sample.odt")) == "odf"
    assert cli_module._detect_validator_for_path(Path("sample.ods")) == "odf"
    assert cli_module._detect_validator_for_path(Path("sample.odp")) == "odf"


def test_cli_auto_routes_odf_to_odf_validator(monkeypatch, minimal_odt: Path) -> None:
    calls: list[tuple[str, str, FileFormat, int, bool]] = []

    class DummyOdfValidator:
        def __init__(self, file_format: FileFormat, max_errors: int, strict: bool) -> None:
            self.file_format = file_format
            self.max_errors = max_errors
            self.strict = strict

        def validate(self, path: Path) -> ValidationResult:
            calls.append(("odf", str(path), self.file_format, self.max_errors, self.strict))
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


def test_cli_auto_routes_ooxml_to_ooxml_validator(monkeypatch, minimal_pptx: Path) -> None:
    calls: list[tuple[str, str, FileFormat, bool]] = []

    class DummyOpenXmlValidator:
        def __init__(self, file_format: FileFormat, max_errors: int, strict: bool) -> None:
            self.file_format = file_format
            self.strict = strict

        def validate(self, path: Path) -> ValidationResult:
            calls.append(("ooxml", str(path), self.file_format, self.strict))
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
