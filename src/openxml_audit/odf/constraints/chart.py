"""Chart validation constraints (M6).

Rules for ODF chart elements: data-range references, axis/series
consistency, and chart style references.
"""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import CHART_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule
from openxml_audit.odf.constraints.style import collect_all_style_names


def _find_chart_roots(ctx: EvaluationContext) -> list[tuple[str, etree._Element]]:
    """Find all chart:chart elements across parsed parts (cached)."""
    return ctx.cached("chart_roots", lambda: _scan_chart_roots(ctx))  # type: ignore[return-value]


def _scan_chart_roots(ctx: EvaluationContext) -> list[tuple[str, etree._Element]]:
    """Scan all parsed parts for chart:chart elements."""
    results: list[tuple[str, etree._Element]] = []
    for part_name, root in ctx.parsed_parts.items():
        for chart in root.iter(f"{{{CHART_NS}}}chart"):
            results.append((part_name, chart))
    return results


class ChartPlotAreaConstraint(OdfConstraint):
    """ODFSEMCHART001: chart:chart must contain a chart:plot-area."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHART001",
            family="chart",
            description="Charts must contain a chart:plot-area element.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part_name, chart in _find_chart_roots(ctx):
            plot_area = chart.find(f"{{{CHART_NS}}}plot-area")
            if plot_area is None:
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHART001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="chart:chart is missing required chart:plot-area",
                        part_uri=self._normalize_part_uri(part_name),
                    )
                )
        return errors


class ChartAxisConstraint(OdfConstraint):
    """ODFSEMCHART002: chart:plot-area should have at least one chart:axis."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHART002",
            family="chart",
            description="Plot areas should have at least one chart:axis.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part_name, chart in _find_chart_roots(ctx):
            plot_area = chart.find(f"{{{CHART_NS}}}plot-area")
            if plot_area is None:
                continue
            axes = plot_area.findall(f"{{{CHART_NS}}}axis")
            if not axes:
                errors.append(
                    self._error(
                        rule_id="ODFSEMCHART002",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="chart:plot-area has no chart:axis elements",
                        part_uri=self._normalize_part_uri(part_name),
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors


class ChartSeriesDataRangeConstraint(OdfConstraint):
    """ODFSEMCHART003: chart:series cell-range-address must not be empty."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHART003",
            family="chart",
            description="Chart series cell range addresses must not be empty.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part_name, chart in _find_chart_roots(ctx):
            plot_area = chart.find(f"{{{CHART_NS}}}plot-area")
            if plot_area is None:
                continue
            for series in plot_area.iter(f"{{{CHART_NS}}}series"):
                range_addr = series.get(
                    f"{{{CHART_NS}}}values-cell-range-address", ""
                ).strip()
                # If the attribute is present but empty, that's an error
                if (
                    series.get(f"{{{CHART_NS}}}values-cell-range-address") is not None
                    and not range_addr
                ):
                    errors.append(
                        self._error(
                            rule_id="ODFSEMCHART003",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                "chart:series has empty "
                                "chart:values-cell-range-address attribute"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors


class ChartStyleRefConstraint(OdfConstraint):
    """ODFSEMCHART004: Chart element style references must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMCHART004",
            family="chart",
            description="Chart element style:name references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part_name, chart in _find_chart_roots(ctx):
            style_names: set[str] = set()
            part_root = ctx.parsed_parts.get(part_name)
            styles_root = ctx.parsed_parts.get("styles.xml")
            if part_root is not None:
                style_names = collect_all_style_names(part_root, styles_root)

            # Check chart elements for style-name references
            reported: set[str] = set()
            for elem in chart.iter():
                if not isinstance(elem.tag, str):
                    continue
                qname = etree.QName(elem)
                if qname.namespace != CHART_NS:
                    continue
                ref = elem.get(f"{{{CHART_NS}}}style-name", "").strip()
                if ref and ref not in style_names and ref not in reported:
                    reported.add(ref)
                    errors.append(
                        self._error(
                            rule_id="ODFSEMCHART004",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Chart element '{qname.localname}' references "
                                f"style '{ref}' which is not defined"
                            ),
                            part_uri=self._normalize_part_uri(part_name),
                        )
                    )
        return errors
