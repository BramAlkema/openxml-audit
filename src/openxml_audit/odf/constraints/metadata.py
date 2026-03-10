"""Metadata-related ODF constraints."""

from __future__ import annotations

from openxml_audit.errors import ValidationError, ValidationErrorType
from openxml_audit.odf._helpers import META_NS, OFFICE_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule


class MetaStatisticsConstraint(OdfConstraint):
    """ODFSEMMETA001: document-statistic attributes must be non-negative integers."""

    COUNT_ATTRS = [
        f"{{{META_NS}}}page-count",
        f"{{{META_NS}}}table-count",
        f"{{{META_NS}}}image-count",
        f"{{{META_NS}}}object-count",
        f"{{{META_NS}}}paragraph-count",
        f"{{{META_NS}}}word-count",
        f"{{{META_NS}}}character-count",
        f"{{{META_NS}}}non-whitespace-character-count",
        f"{{{META_NS}}}cell-count",
        f"{{{META_NS}}}sentence-count",
        f"{{{META_NS}}}syllable-count",
        f"{{{META_NS}}}row-count",
    ]

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMMETA001",
            family="metadata",
            description="Document statistics attributes must be non-negative integers.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        meta = ctx.parsed_parts.get("meta.xml")
        if meta is None:
            return errors

        office_meta = meta.find(f"{{{OFFICE_NS}}}meta")
        if office_meta is None:
            return errors

        stats = office_meta.find(f"{{{META_NS}}}document-statistic")
        if stats is None:
            return errors

        for attr in self.COUNT_ATTRS:
            value = stats.get(attr)
            if value is None:
                continue
            value = value.strip()
            try:
                n = int(value)
                if n < 0:
                    raise ValueError("negative")
            except ValueError:
                local = attr.split("}")[-1]
                errors.append(
                    self._error(
                        rule_id="ODFSEMMETA001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Document statistic '{local}' has invalid value "
                            f"'{value}' (expected non-negative integer)"
                        ),
                        part_uri="/meta.xml",
                    )
                )
        return errors
