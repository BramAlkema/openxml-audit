"""Spreadsheet-related ODF constraints."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import OFFICE_NS, STYLE_NS, TABLE_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule
from openxml_audit.odf.constraints.style import collect_all_style_names


class SpreadsheetTableNameConstraint(OdfConstraint):
    """ODFSEMSS001: Spreadsheet table names must be present and unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS001",
            family="spreadsheet",
            description="Spreadsheet table names must be present and unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            if not name:
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description="Spreadsheet table is missing required table:name",
                        part_uri="/content.xml",
                    )
                )
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS001",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate spreadsheet table name '{name}'",
                        part_uri="/content.xml",
                    )
                )
                continue
            seen.add(name)
        return errors


class SpreadsheetNamedRangeConstraint(OdfConstraint):
    """ODFSEMSS002: Named ranges must reference valid table names."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS002",
            family="spreadsheet",
            description="Named ranges must reference valid table names.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        table_names: set[str] = set()
        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            if name:
                table_names.add(name)

        for nr in content.xpath(".//table:named-range", namespaces={"table": TABLE_NS}):
            base_cell = nr.get(f"{{{TABLE_NS}}}base-cell-address", "").strip()
            if not base_cell:
                continue
            table_ref = base_cell.split(".")[0].strip("$").strip("'")
            if table_ref and table_ref not in table_names:
                range_name = nr.get(f"{{{TABLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS002",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Named range '{range_name or '(unnamed)'}' references "
                            f"table '{table_ref}' which does not exist"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class SpreadsheetColumnCountConstraint(OdfConstraint):
    """ODFSEMSS003: Row cell count must not exceed declared column count."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS003",
            family="spreadsheet",
            description="Column count in rows must not exceed table column definition.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            table_name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            col_count = 0
            for col in table.iterchildren(f"{{{TABLE_NS}}}table-column"):
                repeat = col.get(
                    f"{{{TABLE_NS}}}number-columns-repeated", "1"
                ).strip()
                try:
                    col_count += int(repeat)
                except ValueError:
                    col_count += 1

            if col_count == 0:
                continue

            for row in table.iterchildren(f"{{{TABLE_NS}}}table-row"):
                cell_count = 0
                for cell in row.iterchildren(
                    f"{{{TABLE_NS}}}table-cell",
                    f"{{{TABLE_NS}}}covered-table-cell",
                ):
                    repeat = cell.get(
                        f"{{{TABLE_NS}}}number-columns-repeated", "1"
                    ).strip()
                    try:
                        cell_count += int(repeat)
                    except ValueError:
                        cell_count += 1

                if cell_count > col_count:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSS003",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Table '{table_name or '(unnamed)'}' has a row with "
                                f"{cell_count} cells but only {col_count} columns defined"
                            ),
                            part_uri="/content.xml",
                            severity=ValidationSeverity.WARNING,
                        )
                    )
                    break
        return errors


# ── New M2 rules ────────────────────────────────────────────────────────


class SpreadsheetMinTableConstraint(OdfConstraint):
    """ODFSEMSS004: Spreadsheet must have at least one table."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS004",
            family="spreadsheet",
            description="Spreadsheet must contain at least one table.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return []

        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            return []
        ss = body.find(f"{{{OFFICE_NS}}}spreadsheet")
        if ss is None:
            return []

        tables = ss.findall(f"{{{TABLE_NS}}}table")
        if tables:
            return []

        return [
            self._error(
                rule_id="ODFSEMSS004",
                error_type=ValidationErrorType.SEMANTIC,
                description="Spreadsheet body contains no table:table elements",
                part_uri="/content.xml",
            )
        ]


class DatabaseRangeTableConstraint(OdfConstraint):
    """ODFSEMSS005: Database range target table must exist."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS005",
            family="spreadsheet",
            description="Database range target table must exist.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        table_names: set[str] = set()
        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            if name:
                table_names.add(name)

        for dbr in content.iter(f"{{{TABLE_NS}}}database-range"):
            target = dbr.get(f"{{{TABLE_NS}}}target-range-address", "").strip()
            if not target:
                continue
            table_ref = target.split(".")[0].strip("$").strip("'")
            if table_ref and table_ref not in table_names:
                name = dbr.get(f"{{{TABLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS005",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Database range '{name or '(unnamed)'}' references "
                            f"table '{table_ref}' which does not exist"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class DataPilotSourceConstraint(OdfConstraint):
    """ODFSEMSS006: Data pilot source table must exist."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS006",
            family="spreadsheet",
            description="Data pilot source table must exist.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        table_names: set[str] = set()
        for table in content.xpath(".//table:table", namespaces={"table": TABLE_NS}):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            if name:
                table_names.add(name)

        for dp in content.iter(f"{{{TABLE_NS}}}data-pilot-table"):
            src_range = dp.get(f"{{{TABLE_NS}}}source-cell-range-addresses", "").strip()
            if not src_range:
                continue
            table_ref = src_range.split(".")[0].strip("$").strip("'")
            if table_ref and table_ref not in table_names:
                dp_name = dp.get(f"{{{TABLE_NS}}}name", "").strip()
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS006",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Data pilot '{dp_name or '(unnamed)'}' references "
                            f"source table '{table_ref}' which does not exist"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class CellValidationUniqueConstraint(OdfConstraint):
    """ODFSEMSS007: Content validation names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS007",
            family="spreadsheet",
            description="Content validation names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for cv in content.iter(f"{{{TABLE_NS}}}content-validation"):
            name = cv.get(f"{{{TABLE_NS}}}name", "").strip()
            if not name:
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS007",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate content validation name '{name}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(name)
        return errors


class CellValidationRefConstraint(OdfConstraint):
    """ODFSEMSS008: Cell content-validation-name must resolve to a defined validation."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS008",
            family="spreadsheet",
            description="Cell validation references must resolve to definitions.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        defined: set[str] = set()
        for cv in content.iter(f"{{{TABLE_NS}}}content-validation"):
            name = cv.get(f"{{{TABLE_NS}}}name", "").strip()
            if name:
                defined.add(name)

        if not defined:
            return errors

        reported: set[str] = set()
        for cell in content.iter(f"{{{TABLE_NS}}}table-cell"):
            ref = cell.get(f"{{{TABLE_NS}}}content-validation-name", "").strip()
            if not ref or ref in defined or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMSS008",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Cell references validation '{ref}' which is not defined"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors


class RepeatCountConstraint(OdfConstraint):
    """ODFSEMSS009: Row/column repeat counts must be positive integers."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS009",
            family="spreadsheet",
            description="Repeat counts must be positive integers.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        repeat_attrs = (
            (f"{{{TABLE_NS}}}table-row", f"{{{TABLE_NS}}}number-rows-repeated"),
            (f"{{{TABLE_NS}}}table-column", f"{{{TABLE_NS}}}number-columns-repeated"),
            (f"{{{TABLE_NS}}}table-cell", f"{{{TABLE_NS}}}number-columns-repeated"),
        )

        reported = False
        for tag, attr in repeat_attrs:
            for elem in content.iter(tag):
                val = elem.get(attr, "").strip()
                if not val:
                    continue
                try:
                    n = int(val)
                    if n < 1:
                        raise ValueError("non-positive")
                except ValueError:
                    errors.append(
                        self._error(
                            rule_id="ODFSEMSS009",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(
                                f"Invalid repeat count '{val}' on "
                                f"{etree.QName(elem).localname}"
                            ),
                            part_uri="/content.xml",
                        )
                    )
                    reported = True
                    break
            if reported:
                break
        return errors


class ColumnStyleRefConstraint(OdfConstraint):
    """ODFSEMSS010: Column default-cell-style-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS010",
            family="spreadsheet",
            description="Column default cell style references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        styles = ctx.parsed_parts.get("styles.xml")
        all_names = collect_all_style_names(content, styles)
        if not all_names:
            return errors

        reported: set[str] = set()
        for col in content.iter(f"{{{TABLE_NS}}}table-column"):
            ref = col.get(f"{{{TABLE_NS}}}default-cell-style-name", "").strip()
            if not ref or ref in all_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMSS010",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Column default cell style '{ref}' is not defined"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors


class CellStyleRefConstraint(OdfConstraint):
    """ODFSEMSS011: table:table-cell style-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS011",
            family="spreadsheet",
            description="Cell style references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        styles = ctx.parsed_parts.get("styles.xml")
        all_names = collect_all_style_names(content, styles)
        if not all_names:
            return errors

        reported: set[str] = set()
        for cell in content.iter(f"{{{TABLE_NS}}}table-cell"):
            ref = cell.get(f"{{{TABLE_NS}}}style-name", "").strip()
            if not ref or ref in all_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMSS011",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=f"Cell style '{ref}' is not defined",
                    part_uri="/content.xml",
                )
            )
        return errors


class ConditionalStyleRefConstraint(OdfConstraint):
    """ODFSEMSS012: Conditional style apply-style-name must resolve."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS012",
            family="spreadsheet",
            description="Conditional format style references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        styles_root = ctx.parsed_parts.get("styles.xml")
        all_names = collect_all_style_names(content, styles_root)

        reported: set[str] = set()
        for sm in content.iter(f"{{{STYLE_NS}}}map"):
            ref = sm.get(f"{{{STYLE_NS}}}apply-style-name", "").strip()
            if not ref or ref in all_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMSS012",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Conditional style '{ref}' is not defined"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors


class FilterFieldConstraint(OdfConstraint):
    """ODFSEMSS013: Filter field-number must be a non-negative integer."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMSS013",
            family="spreadsheet",
            description="Filter field numbers must be non-negative integers.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.spreadsheet"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        for cond in content.iter(f"{{{TABLE_NS}}}filter-condition"):
            field = cond.get(f"{{{TABLE_NS}}}field-number", "").strip()
            if not field:
                continue
            try:
                n = int(field)
                if n < 0:
                    raise ValueError("negative")
            except ValueError:
                errors.append(
                    self._error(
                        rule_id="ODFSEMSS013",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Filter field-number '{field}' is not a "
                            "valid non-negative integer"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors
