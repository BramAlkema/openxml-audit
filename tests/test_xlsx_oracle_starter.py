"""XLSX calibration-emitter tests: built packages must validate."""

from __future__ import annotations

from pathlib import Path

from openxml_audit import OpenXmlValidator
from openxml_audit.xlsx.oracle_starter_book import build_minimal_xlsx


def test_minimal_xlsx_has_valid_package_structure(tmp_path: Path) -> None:
    """Package tier: ZIP opens, content-types parse, rels resolve, parts exist.

    Schema + semantic checks are skipped — a calibration probe is stripped
    to one feature on purpose. Excel's conventions (required styles.xml and
    theme.xml relationships) are exactly what the osa re-save reveals as
    the tier-escalation diff.
    """
    output = tmp_path / "minimal.xlsx"
    build_minimal_xlsx(output)

    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
    result = validator.validate(output)

    assert result.is_valid, (
        f"Minimal XLSX failed package validation: "
        f"{[f'{e.severity.value}: {e.description}' for e in result.errors]}"
    )
