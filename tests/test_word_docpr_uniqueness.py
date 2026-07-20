"""Tests for cross-part wp:docPr id uniqueness (issue #6).

Word renumbers a drawing whose `wp:docPr/@id` collides with one in another
part of the same document story, so the package is repair-worthy even though
the SDK's per-part schematron accepts it.
"""

from __future__ import annotations

from pathlib import Path

from openxml_audit import OpenXmlValidator
from openxml_audit.errors import SourceClass, ValidationSeverity
from openxml_audit.verdict import AppPrediction, predict

ERROR_ID = "Sem_DuplicateDocPrId"


def _docpr_findings(path: Path) -> list:
    result = OpenXmlValidator().validate(str(path))
    return [e for e in result.errors if e.id == ERROR_ID]


def test_duplicate_docpr_across_document_and_footer_is_reported(
    docx_duplicate_docpr: Path,
) -> None:
    findings = _docpr_findings(docx_duplicate_docpr)

    assert len(findings) == 1
    finding = findings[0]
    assert finding.part_uri == "/word/footer2.xml"
    assert "296" in finding.description
    assert "/word/document.xml" in finding.description


def test_distinct_docpr_ids_are_not_reported(
    docx_duplicate_docpr_unique: Path,
) -> None:
    assert _docpr_findings(docx_duplicate_docpr_unique) == []


def test_finding_is_app_compat_not_sdk_parity(
    docx_duplicate_docpr: Path,
) -> None:
    # The SDK's schematron for this rule is scoped to a single part, so it
    # does not report the cross-part case. Classifying the finding as
    # SDK_PROXY would make the parity gate expect a divergence that is not
    # actually there.
    (finding,) = _docpr_findings(docx_duplicate_docpr)
    assert finding.source_class is SourceClass.WORD_APP_COMPAT


def test_reported_against_the_later_part_naming_the_first(
    docx_duplicate_docpr: Path,
) -> None:
    # The main document is scanned before its headers/footers, so the
    # collision is always attributed to the header/footer that repeats an id
    # the document already used — matching how Word renumbers the footer.
    (finding,) = _docpr_findings(docx_duplicate_docpr)
    expected = "duplicate wp:docPr id 296; first seen in /word/document.xml"
    assert finding.description == expected


def test_clean_document_has_no_docpr_findings(minimal_docx: Path) -> None:
    assert _docpr_findings(minimal_docx) == []


def test_collision_is_a_warning_not_an_error(docx_duplicate_docpr: Path) -> None:
    # Word repairs the file rather than refusing it, so ERROR would
    # over-promise — same call `word.compat` makes for its ordering rules.
    (finding,) = _docpr_findings(docx_duplicate_docpr)
    assert finding.severity is ValidationSeverity.WARNING


def test_collision_still_predicts_repair(docx_duplicate_docpr: Path) -> None:
    # Warning severity must not soften the survival verdict: the verdict
    # layer keys off the app-compat source class at any severity.
    result = OpenXmlValidator().validate(str(docx_duplicate_docpr))
    assert predict(result).prediction is AppPrediction.REPAIR_OR_REWRITE_LIKELY


def test_distinct_ids_predict_no_repair(docx_duplicate_docpr_unique: Path) -> None:
    result = OpenXmlValidator().validate(str(docx_duplicate_docpr_unique))
    assert predict(result).prediction is not AppPrediction.REPAIR_OR_REWRITE_LIKELY
