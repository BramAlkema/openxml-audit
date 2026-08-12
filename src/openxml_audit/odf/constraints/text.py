"""Text-related ODF constraints."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import OFFICE_NS, TABLE_NS, TEXT_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule
from openxml_audit.odf.constraints.style import collect_list_style_names


class TextStyleReferenceConstraint(OdfConstraint):
    """ODFSEMTXT001: Text style references require styles.xml companion part."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT001",
            family="text",
            description="Text style references require styles.xml companion part.",
        )

    @staticmethod
    def _has_text_style_references(content: etree._Element) -> bool:
        style_attr = f"{{{TEXT_NS}}}style-name"
        for elem in content.iter():
            tag = elem.tag
            if not isinstance(tag, str) or not tag.startswith(f"{{{TEXT_NS}}}"):
                continue
            value = elem.get(style_attr)
            if value is not None and value.strip():
                return True
        return False

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return []
        if "styles.xml" in ctx.package.manifest_paths():
            return []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return []
        if not self._has_text_style_references(content):
            return []
        return [
            self._error(
                rule_id="ODFSEMTXT001",
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    "content.xml contains text:style-name references but styles.xml "
                    "is not declared in manifest.xml"
                ),
                part_uri="/content.xml",
            )
        ]


class TextListLevelConstraint(OdfConstraint):
    """ODFSEMTXT002: text:list must reference defined list styles."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT002",
            family="text",
            description="List level style references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        styles = ctx.parsed_parts.get("styles.xml")
        list_names = collect_list_style_names(content, styles)

        reported: set[str] = set()
        for lst in content.iter(f"{{{TEXT_NS}}}list"):
            ref = lst.get(f"{{{TEXT_NS}}}style-name", "").strip()
            if not ref or ref in list_names or ref in reported:
                continue
            reported.add(ref)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT002",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(f"text:list references list style '{ref}' which is not defined"),
                    part_uri="/content.xml",
                )
            )
        return errors


class TextBookmarkRefConstraint(OdfConstraint):
    """ODFSEMTXT003: bookmark-ref must point to defined bookmarks."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT003",
            family="text",
            description="Bookmark references must resolve to defined bookmarks.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        bookmark_names: set[str] = set()
        for elem in content.iter(f"{{{TEXT_NS}}}bookmark", f"{{{TEXT_NS}}}bookmark-start"):
            name = elem.get(f"{{{TEXT_NS}}}name", "").strip()
            if name:
                bookmark_names.add(name)

        reported: set[str] = set()
        for ref_elem in content.iter(f"{{{TEXT_NS}}}bookmark-ref"):
            ref_name = ref_elem.get(f"{{{TEXT_NS}}}ref-name", "").strip()
            if not ref_name or ref_name in bookmark_names or ref_name in reported:
                continue
            reported.add(ref_name)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT003",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"Bookmark reference '{ref_name}' does not resolve to a defined bookmark"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors


# ── New M2 rules ────────────────────────────────────────────────────────


class HeadingLevelSkipConstraint(OdfConstraint):
    """ODFSEMTXT004: Heading outline levels must not skip levels."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT004",
            family="text",
            description="Heading outline levels must not skip levels.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        prev_level = 0
        for h in content.iter(f"{{{TEXT_NS}}}h"):
            level_str = h.get(f"{{{TEXT_NS}}}outline-level", "1").strip()
            try:
                level = int(level_str)
            except ValueError:
                continue
            if level < 1:
                continue
            if prev_level > 0 and level > prev_level + 1:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(f"Heading skips from level {prev_level} to {level}"),
                        part_uri="/content.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
                break  # one warning is enough
            prev_level = level
        return errors


class NoteRefConstraint(OdfConstraint):
    """ODFSEMTXT005: text:note-ref must reference a defined note."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT005",
            family="text",
            description="Note references must resolve to defined notes.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        note_ids: set[str] = set()
        for note in content.iter(f"{{{TEXT_NS}}}note"):
            nid = note.get(f"{{{TEXT_NS}}}id", "").strip()
            if nid:
                note_ids.add(nid)

        reported: set[str] = set()
        for ref in content.iter(f"{{{TEXT_NS}}}note-ref"):
            ref_id = ref.get(f"{{{TEXT_NS}}}ref-name", "").strip()
            if not ref_id or ref_id in note_ids or ref_id in reported:
                continue
            reported.add(ref_id)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT005",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(f"Note reference '{ref_id}' does not resolve to a defined note"),
                    part_uri="/content.xml",
                )
            )
        return errors


class SectionNameUniqueConstraint(OdfConstraint):
    """ODFSEMTXT006: Section names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT006",
            family="text",
            description="Section names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for section in content.iter(f"{{{TEXT_NS}}}section"):
            name = section.get(f"{{{TEXT_NS}}}name", "").strip()
            if not name:
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT006",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate section name '{name}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(name)
        return errors


class TrackedChangeIdConstraint(OdfConstraint):
    """ODFSEMTXT007: Tracked change IDs must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT007",
            family="text",
            description="Tracked change IDs must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for change in content.iter(f"{{{TEXT_NS}}}changed-region"):
            cid = change.get(f"{{{TEXT_NS}}}id", "").strip()
            if not cid:
                continue
            if cid in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT007",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate tracked change ID '{cid}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(cid)
        return errors


class SequenceDeclUniqueConstraint(OdfConstraint):
    """ODFSEMTXT008: text:sequence-decl names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT008",
            family="text",
            description="Sequence declaration names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for decl in content.iter(f"{{{TEXT_NS}}}sequence-decl"):
            name = decl.get(f"{{{TEXT_NS}}}name", "").strip()
            if not name:
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT008",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate sequence declaration name '{name}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(name)
        return errors


class VariableDeclUniqueConstraint(OdfConstraint):
    """ODFSEMTXT009: text:variable-decl names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT009",
            family="text",
            description="Variable declaration names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for decl in content.iter(f"{{{TEXT_NS}}}variable-decl"):
            name = decl.get(f"{{{TEXT_NS}}}name", "").strip()
            if not name:
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT009",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate variable declaration '{name}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(name)
        return errors


class VariableGetRefConstraint(OdfConstraint):
    """ODFSEMTXT010: text:variable-get must reference a declared variable."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT010",
            family="text",
            description="Variable get references must resolve to declarations.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        declared: set[str] = set()
        for decl in content.iter(
            f"{{{TEXT_NS}}}variable-decl",
            f"{{{TEXT_NS}}}variable-set",
        ):
            name = decl.get(f"{{{TEXT_NS}}}name", "").strip()
            if name:
                declared.add(name)

        reported: set[str] = set()
        for ref in content.iter(f"{{{TEXT_NS}}}variable-get"):
            name = ref.get(f"{{{TEXT_NS}}}name", "").strip()
            if not name or name in declared or name in reported:
                continue
            reported.add(name)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT010",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(f"Variable reference '{name}' does not resolve to a declaration"),
                    part_uri="/content.xml",
                )
            )
        return errors


class UserFieldDeclUniqueConstraint(OdfConstraint):
    """ODFSEMTXT011: text:user-field-decl names must be unique."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT011",
            family="text",
            description="User field declaration names must be unique.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        seen: set[str] = set()
        for decl in content.iter(f"{{{TEXT_NS}}}user-field-decl"):
            name = decl.get(f"{{{TEXT_NS}}}name", "").strip()
            if not name:
                continue
            if name in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT011",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=f"Duplicate user field declaration '{name}'",
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(name)
        return errors


class UserFieldGetRefConstraint(OdfConstraint):
    """ODFSEMTXT012: text:user-field-get must reference a declared user field."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT012",
            family="text",
            description="User field references must resolve to declarations.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        declared: set[str] = set()
        for decl in content.iter(f"{{{TEXT_NS}}}user-field-decl"):
            name = decl.get(f"{{{TEXT_NS}}}name", "").strip()
            if name:
                declared.add(name)

        reported: set[str] = set()
        for ref in content.iter(f"{{{TEXT_NS}}}user-field-get"):
            name = ref.get(f"{{{TEXT_NS}}}name", "").strip()
            if not name or name in declared or name in reported:
                continue
            reported.add(name)
            errors.append(
                self._error(
                    rule_id="ODFSEMTXT012",
                    error_type=ValidationErrorType.SEMANTIC,
                    description=(
                        f"User field reference '{name}' does not resolve to a declaration"
                    ),
                    part_uri="/content.xml",
                )
            )
        return errors


class TextTableStructureConstraint(OdfConstraint):
    """ODFSEMTXT013: Tables in text documents must have at least one column and one row."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMTXT013",
            family="text",
            description="Tables in text documents must have columns and rows.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        mimetype = ctx.package.mimetype or ""
        if not mimetype.startswith("application/vnd.oasis.opendocument.text"):
            return errors
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        body = content.find(f"{{{OFFICE_NS}}}body")
        if body is None:
            return errors
        text_body = body.find(f"{{{OFFICE_NS}}}text")
        if text_body is None:
            return errors

        for table in text_body.iter(f"{{{TABLE_NS}}}table"):
            name = table.get(f"{{{TABLE_NS}}}name", "").strip()
            cols = list(table.iterchildren(f"{{{TABLE_NS}}}table-column"))
            rows = list(table.iterchildren(f"{{{TABLE_NS}}}table-row"))
            if not cols and not rows:
                errors.append(
                    self._error(
                        rule_id="ODFSEMTXT013",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(f"Table '{name or '(unnamed)'}' has no columns or rows"),
                        part_uri="/content.xml",
                    )
                )
        return errors
