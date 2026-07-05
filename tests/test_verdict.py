"""Tests for per-app survival verdicts (Spec 035)."""

from __future__ import annotations

import json
from pathlib import Path

from openxml_audit import cli as cli_module
from openxml_audit.errors import (
    FileFormat,
    SourceClass,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.validator import OpenXmlValidator
from openxml_audit.verdict import AppPrediction, predict, target_app

TOKENMOULDS_XLSX = Path("data/corpus/tokenmoulds_v0.7.2/excel/acme-us.xlsx")


def _finding(
    description: str,
    *,
    severity: ValidationSeverity = ValidationSeverity.ERROR,
    source_class: SourceClass = SourceClass.SDK_PROXY,
) -> ValidationError:
    return ValidationError(
        error_type=ValidationErrorType.SEMANTIC,
        severity=severity,
        description=description,
        part_uri="/part.xml",
        source_class=source_class,
    )


def _result(file_path: str, *errors: ValidationError) -> ValidationResult:
    return ValidationResult(
        is_valid=not any(e.severity == ValidationSeverity.ERROR for e in errors),
        errors=list(errors),
        file_path=file_path,
        file_format=FileFormat.OFFICE_2019,
    )


class TestTargetApp:
    def test_extension_mapping(self):
        assert target_app("a.docx") == "Microsoft Word"
        assert target_app("a.XLSM") == "Microsoft Excel"
        assert target_app("a.ppsx") == "Microsoft PowerPoint"
        assert target_app("a.odt") == "LibreOffice Writer"
        assert target_app("a.ods") == "LibreOffice Calc"
        assert target_app("a.odp") == "LibreOffice Impress"

    def test_unknown_extension_falls_back(self):
        assert target_app("a.bin") == "the target application"


class TestPredictLadder:
    def test_clean_file_opens_clean(self):
        verdict = predict(_result("deck.pptx"))
        assert verdict.prediction is AppPrediction.OPENS_CLEAN
        assert verdict.headline == ("Microsoft PowerPoint: expected to open cleanly (no findings)")
        assert verdict.basis == ()

    def test_non_compat_warnings_stay_opens_clean(self):
        verdict = predict(
            _result(
                "doc.docx",
                _finding("external rel", severity=ValidationSeverity.WARNING),
            )
        )
        assert verdict.prediction is AppPrediction.OPENS_CLEAN
        assert "1 warning only" in verdict.headline

    def test_package_integrity_wins_over_everything(self):
        verdict = predict(
            _result(
                "doc.docx",
                _finding("broken rels", source_class=SourceClass.PACKAGE_INTEGRITY),
                _finding("word compat", source_class=SourceClass.WORD_APP_COMPAT),
            )
        )
        assert verdict.prediction is AppPrediction.REJECT_LIKELY
        assert "unlikely to open" in verdict.headline
        assert verdict.basis == ("broken rels",)

    def test_app_compat_predicts_repair_at_any_severity(self):
        verdict = predict(
            _result(
                "book.xlsx",
                _finding(
                    "Excel will rewrite inline strings",
                    severity=ValidationSeverity.WARNING,
                    source_class=SourceClass.EXCEL_APP_COMPAT,
                ),
            )
        )
        assert verdict.prediction is AppPrediction.REPAIR_OR_REWRITE_LIKELY
        assert "repair or rewrite" in verdict.headline

    def test_other_apps_compat_findings_do_not_claim_repair(self):
        # A Word app-compat finding on a .pptx is a generic error signal
        # for PowerPoint, not a PowerPoint repair prediction.
        verdict = predict(
            _result(
                "deck.pptx",
                _finding("word compat", source_class=SourceClass.WORD_APP_COMPAT),
            )
        )
        assert verdict.prediction is AppPrediction.AT_RISK

    def test_schema_errors_are_at_risk(self):
        verdict = predict(_result("deck.pptx", _finding("bad element")))
        assert verdict.prediction is AppPrediction.AT_RISK
        assert "may open, repair, or reject" in verdict.headline

    def test_basis_truncates_after_three(self):
        verdict = predict(_result("deck.pptx", *[_finding(f"e{i}") for i in range(5)]))
        assert verdict.basis == ("e0", "e1", "e2", "... and 2 more")


class TestVerdictAgainstOracleBaseline:
    def test_tokenmoulds_xlsx_matches_observed_excel_rewrite(self):
        # The v0.7.2 oracle baseline observed Excel rewriting this exact
        # workbook on save; the verdict must say so.
        result = OpenXmlValidator().validate(TOKENMOULDS_XLSX)
        verdict = predict(result)
        assert verdict.prediction is AppPrediction.REPAIR_OR_REWRITE_LIKELY
        assert verdict.app == "Microsoft Excel"


class TestCliOutput:
    def test_text_output_leads_with_headline(self):
        result = _result("deck.pptx", _finding("bad element"))
        with cli_module.console.capture() as capture:
            cli_module._output_text([result], quiet=False)
        output = capture.get()
        first_line = output.splitlines()[0]
        assert first_line.startswith("Microsoft PowerPoint: at risk")
        assert "1. bad element" in output

    def test_json_output_carries_verdict_block(self):
        result = _result("book.xlsx")
        with cli_module.console.capture() as capture:
            cli_module._output_json([result])
        payload = json.loads(capture.get())
        verdict = payload[0]["verdict"]
        assert verdict["prediction"] == "opens-clean"
        assert verdict["app"] == "Microsoft Excel"
        assert verdict["basis"] == []
