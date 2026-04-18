"""DOCX calibration-emitter tests: built packages must validate."""

from __future__ import annotations

from pathlib import Path

from openxml_audit import OpenXmlValidator
from openxml_audit.docx.oracle_starter_doc import build_minimal_docx


def test_minimal_docx_has_valid_package_structure(tmp_path: Path) -> None:
    """Package tier: ZIP opens, content-types parse, rels resolve, parts exist.

    Schema + semantic checks are skipped — a calibration probe is stripped
    to one feature on purpose. Word's conventions (required styles.xml,
    settings.xml, theme.xml relationships) are exactly what the osa
    re-save reveals as the tier-escalation diff.
    """
    output = tmp_path / "minimal.docx"
    build_minimal_docx(output)

    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(output)

    assert result.is_valid, (
        f"Minimal DOCX failed package validation: "
        f"{[f'{e.severity.value}: {e.description}' for e in result.errors]}"
    )
