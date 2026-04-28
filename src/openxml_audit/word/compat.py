"""Word compatibility checks beyond ECMA-376 schema validation.

Implements the Word-app-compat ordering check described in spec 010:
detects child element reorderings inside WordprocessingML property
complex types that trigger Word's "unreadable content" repair dialog
despite the .NET Open XML SDK considering the same files valid.

Phase 1 scope: CT_TrPr (table row properties).

The canonical child ordering is sourced from the SDK schema metadata
(`Children` field for the matching type), which preserves the XSD's
declarative order even though the SDK's runtime particle relaxes it.
Re-derive when ECMA-376 revises a property type — see
`data/openxml/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json`.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.errors import (
    ValidationError,
    ValidationErrorType,
    ValidationSeverity,
)
from openxml_audit.namespaces import WORDPROCESSINGML

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import DocumentPart


W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _w(local: str) -> str:
    return f"{{{WORDPROCESSINGML}}}{local}"


def _w14(local: str) -> str:
    return f"{{{W14_NS}}}{local}"


def _local_name(clark_tag: str) -> str:
    """Strip Clark-notation namespace prefix for human-readable messages."""
    if clark_tag.startswith("{"):
        return clark_tag.split("}", 1)[1]
    return clark_tag


@dataclass(frozen=True)
class ChildSequence:
    """Canonical child element ordering for a WML property complex type.

    Observed children must form a subsequence of `children` — they may skip
    canonical entries (every child is effectively optional in these types),
    but they may not reorder them.
    """

    parent_tag: str  # Clark-notation parent element
    parent_local: str  # short name for error messages
    spec_section: str  # ECMA-376 reference for traceability
    children: tuple[str, ...]  # Clark-notation children in canonical order


# CT_TrPr canonical child order, derived from SDK schema metadata
# `w:CT_TrPr/w:trPr` Children. Word repairs files where these children
# appear in any other order even though the runtime particle accepts them.
_CT_TR_PR = ChildSequence(
    parent_tag=_w("trPr"),
    parent_local="trPr",
    spec_section="ECMA-376 §17.4.79",
    children=(
        _w("cnfStyle"),
        _w("divId"),
        _w("gridBefore"),
        _w("gridAfter"),
        _w("wBefore"),
        _w("wAfter"),
        _w("trHeight"),
        _w("hidden"),
        _w("cantSplit"),
        _w("tblHeader"),
        _w("tblCellSpacing"),
        _w("jc"),
        _w("ins"),
        _w("del"),
        _w("trPrChange"),
        _w14("conflictIns"),
        _w14("conflictDel"),
    ),
)


CONSTRAINT_TABLE: dict[str, ChildSequence] = {
    _CT_TR_PR.parent_tag: _CT_TR_PR,
}


def find_first_out_of_order(
    observed_tags: list[str], canonical: tuple[str, ...]
) -> tuple[int, int] | None:
    """Check whether `observed_tags` is a subsequence of `canonical`.

    Returns `(offending_index, blocking_index)` for the first observed tag
    that breaks the order — i.e. an observed tag whose canonical position is
    earlier than a previously matched observed tag. The blocking index points
    to the previously matched observed tag that already advanced past it.
    Returns None if the sequence is well-ordered.

    Tags that don't appear in `canonical` (e.g., extension elements) are
    silently skipped: ordering checks are scoped to known canonical children.
    """
    canon_index = {tag: i for i, tag in enumerate(canonical)}
    last_canon_pos = -1
    blocking_idx = -1
    for i, tag in enumerate(observed_tags):
        pos = canon_index.get(tag)
        if pos is None:
            continue
        if pos < last_canon_pos:
            return i, blocking_idx
        last_canon_pos = pos
        blocking_idx = i
    return None


class WordCompatValidator:
    """Run Word-app-compat ordering checks over a Word document part."""

    def validate(
        self, part: DocumentPart, context: ValidationContext
    ) -> list[ValidationError]:
        """Walk `part` and emit a WARNING for every property element whose
        children violate the canonical ordering."""
        xml = part.xml
        if xml is None:
            return []

        before = len(context.errors)

        for elem in xml.iter():
            if not isinstance(elem.tag, str):
                continue  # comments, processing instructions
            constraint = CONSTRAINT_TABLE.get(elem.tag)
            if constraint is None:
                continue
            self._check_element(elem, constraint, context)

        return list(context.errors[before:])

    def _check_element(
        self,
        elem: etree._Element,
        constraint: ChildSequence,
        context: ValidationContext,
    ) -> None:
        observed_tags = [
            child.tag for child in elem if isinstance(child.tag, str)
        ]
        result = find_first_out_of_order(observed_tags, constraint.children)
        if result is None:
            return

        offending_idx, blocking_idx = result
        offending = _local_name(observed_tags[offending_idx])
        blocking = (
            _local_name(observed_tags[blocking_idx])
            if blocking_idx >= 0
            else "(start)"
        )
        context.add_error(
            error_type=ValidationErrorType.SEMANTIC,
            description=(
                f"{constraint.parent_local} child '{offending}' appears after "
                f"'{blocking}' but {constraint.spec_section} places it earlier "
                f"— Word may flag this file as unreadable content"
            ),
            node=offending,
            severity=ValidationSeverity.WARNING,
        )
