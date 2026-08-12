"""Form control validation constraints (M6).

Rules for ODF form controls: implementation references, bound cell/range
validation, and event handler target resolution.
"""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import FORM_NS, SCRIPT_NS, XLINK_NS
from openxml_audit.odf.constraints.base import EvaluationContext, OdfConstraint, OdfSemanticRule

# Form control element local names
_CONTROL_LOCAL_NAMES = frozenset(
    {
        "text",
        "textarea",
        "password",
        "file",
        "formatted-text",
        "fixed-text",
        "combobox",
        "listbox",
        "button",
        "image",
        "checkbox",
        "radio",
        "frame",
        "image-frame",
        "hidden",
        "grid",
        "value-range",
        "generic-control",
    }
)


def _iter_form_controls(root: etree._Element) -> list[etree._Element]:
    """Collect all form:* control elements."""
    controls: list[etree._Element] = []
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        qname = etree.QName(elem)
        if qname.namespace == FORM_NS and qname.localname in _CONTROL_LOCAL_NAMES:
            controls.append(elem)
    return controls


def _get_form_controls(ctx: EvaluationContext) -> list[etree._Element]:
    """Cached form control collection from content.xml."""
    content = ctx.parsed_parts.get("content.xml")
    if content is None:
        return []
    return ctx.cached("form_controls", lambda: _iter_form_controls(content))  # type: ignore[return-value]


class FormControlNameUniqueConstraint(OdfConstraint):
    """ODFSEMFORM001: Form control names must be unique within a form."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMFORM001",
            family="forms",
            description="Form control names must be unique within each form.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        for form in content.iter(f"{{{FORM_NS}}}form"):
            seen: set[str] = set()
            for control in _iter_form_controls(form):
                name = control.get(f"{{{FORM_NS}}}name", "").strip()
                if not name:
                    continue
                if name in seen:
                    form_name = form.get(f"{{{FORM_NS}}}name", "(unnamed)")
                    errors.append(
                        self._error(
                            rule_id="ODFSEMFORM001",
                            error_type=ValidationErrorType.SEMANTIC,
                            description=(f"Duplicate control name '{name}' in form '{form_name}'"),
                            part_uri="/content.xml",
                        )
                    )
                else:
                    seen.add(name)
        return errors


class FormControlIdUniqueConstraint(OdfConstraint):
    """ODFSEMFORM002: Form control IDs must be unique across all forms."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMFORM002",
            family="forms",
            description="Form control IDs must be unique across the document.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []

        seen: set[str] = set()
        for control in _get_form_controls(ctx):
            ctrl_id = control.get(f"{{{FORM_NS}}}id", "").strip()
            if not ctrl_id:
                continue
            if ctrl_id in seen:
                errors.append(
                    self._error(
                        rule_id="ODFSEMFORM002",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(f"Duplicate form control ID '{ctrl_id}'"),
                        part_uri="/content.xml",
                    )
                )
            else:
                seen.add(ctrl_id)
        return errors


class FormColumnRefConstraint(OdfConstraint):
    """ODFSEMFORM003: Grid column form:* references must resolve to controls."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMFORM003",
            family="forms",
            description="Form grid column references must resolve.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        # Collect all form control IDs
        control_ids: set[str] = set()
        for control in _get_form_controls(ctx):
            ctrl_id = control.get(f"{{{FORM_NS}}}id", "").strip()
            if ctrl_id:
                control_ids.add(ctrl_id)

        if not control_ids:
            return errors

        # Check grid columns referencing non-existent controls
        reported: set[str] = set()
        for column in content.iter(f"{{{FORM_NS}}}column"):
            ref = column.get(f"{{{FORM_NS}}}control", "").strip()
            if ref and ref not in control_ids and ref not in reported:
                reported.add(ref)
                errors.append(
                    self._error(
                        rule_id="ODFSEMFORM003",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Form grid column references control '{ref}' which is not defined"
                        ),
                        part_uri="/content.xml",
                    )
                )
        return errors


class FormEventListenerConstraint(OdfConstraint):
    """ODFSEMFORM004: Event listener xlink:href must not be empty."""

    @property
    def rule(self) -> OdfSemanticRule:
        return OdfSemanticRule(
            id="ODFSEMFORM004",
            family="forms",
            description="Script event listener hrefs must not be empty.",
        )

    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        content = ctx.parsed_parts.get("content.xml")
        if content is None:
            return errors

        for listener in content.iter(f"{{{SCRIPT_NS}}}event-listener"):
            href = listener.get(f"{{{XLINK_NS}}}href", "")
            macro_name = listener.get(f"{{{SCRIPT_NS}}}macro-name", "")
            event_name = listener.get(f"{{{SCRIPT_NS}}}event-name", "").strip()
            if not href.strip() and not macro_name.strip() and event_name:
                errors.append(
                    self._error(
                        rule_id="ODFSEMFORM004",
                        error_type=ValidationErrorType.SEMANTIC,
                        description=(
                            f"Event listener for '{event_name}' has neither "
                            "xlink:href nor script:macro-name"
                        ),
                        part_uri="/content.xml",
                        severity=ValidationSeverity.WARNING,
                    )
                )
        return errors
