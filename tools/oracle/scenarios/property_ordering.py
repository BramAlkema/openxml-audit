"""Property-element ordering scenarios for the Word roundtrip oracle.

Builds DOCX files using python-docx (so the document has a real Word-
template foundation: full styles, fonts, relationships) and then mutates
the property-element children via lxml to test specific orderings.

Spec: `specs/011-word-roundtrip-oracle.md` Phase 2.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from docx import Document
from lxml import etree
from tools.oracle.diff import slugify_scenario_id

from openxml_audit.namespaces import WORDPROCESSINGML

W = WORDPROCESSINGML
WNS = f"{{{W}}}"


@dataclass(frozen=True)
class ScenarioSpec:
    """One ordering scenario to materialize and roundtrip."""

    id: str
    parent_local: str  # e.g. "trPr"
    input_children: tuple[str, ...]  # local names in the order to test
    description: str


# Minimum-valid attribute set per CT_TrPr child element, keyed on
# attribute local name (without the w: prefix). Empty dict means "no
# attributes required — element can stand alone." Values were chosen
# to be ECMA-376-conformant and inert (no measurable effect on layout).
TRPR_CHILD_ATTRS: dict[str, dict[str, str]] = {
    # ST_DecimalNumber-typed val
    "divId": {"val": "0"},
    "gridBefore": {"val": "0"},
    "gridAfter": {"val": "0"},
    # ST_Cnf — 12-bit conditional-formatting bitstring
    "cnfStyle": {"val": "000000000000"},
    # ST_TblWidth — value with type
    "wBefore": {"w": "0", "type": "dxa"},
    "wAfter": {"w": "0", "type": "dxa"},
    "tblCellSpacing": {"w": "0", "type": "dxa"},
    # ST_JcTable — limited to center/end/start in M365; "center" is a
    # safe choice across versions.
    "jc": {"val": "center"},
    # OnOff and other empty-OK children
    "trHeight": {},
    "hidden": {},
    "cantSplit": {},
    "tblHeader": {},
}


def materialize_trpr_scenario(spec: ScenarioSpec, dest: Path) -> Path:
    """Build a DOCX with a one-row table whose `trPr` contains
    `spec.input_children` in the requested order. Saves to `dest` and
    returns the path.

    Uses python-docx so the document inherits a real Word-blessed
    template (full styles, fonts, font table, relationships). We then
    rewrite the row's trPr via lxml to inject our exact ordering — the
    same approach Shaun used in issue #3. Each child receives its
    minimum-valid attribute set from `TRPR_CHILD_ATTRS` so Word opens
    the file rather than rejecting it as malformed.
    """
    if spec.parent_local != "trPr":
        raise ValueError(
            f"materialize_trpr_scenario expects parent_local='trPr', "
            f"got {spec.parent_local!r}"
        )

    doc = Document()
    table = doc.add_table(rows=1, cols=2)
    tr = table.rows[0]._tr  # noqa: SLF001 — python-docx exposes the lxml element this way

    # Remove any existing trPr
    for existing in tr.findall(f"{WNS}trPr"):
        tr.remove(existing)

    # Build the new trPr with children in the requested order, insert at index 0
    trPr = etree.Element(f"{WNS}trPr")
    tr.insert(0, trPr)
    for child_local in spec.input_children:
        child = etree.SubElement(trPr, f"{WNS}{child_local}")
        for attr_local, attr_val in TRPR_CHILD_ATTRS.get(child_local, {}).items():
            child.set(f"{WNS}{attr_local}", attr_val)

    dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dest))
    return dest


def trpr_baseline(canonical: tuple[str, ...]) -> ScenarioSpec:
    """The canonical-order trPr — the control case. Preserved → engine works."""
    return ScenarioSpec(
        id=slugify_scenario_id("trPr", "baseline"),
        parent_local="trPr",
        input_children=tuple(canonical),
        description="Canonical CT_TrPr child order (control)",
    )


def trpr_pairwise_swaps(canonical: tuple[str, ...]) -> list[ScenarioSpec]:
    """Every pairwise swap among canonical children. For n canonical
    children this produces n*(n-1)/2 scenarios."""
    scenarios: list[ScenarioSpec] = []
    for i in range(len(canonical)):
        for j in range(i + 1, len(canonical)):
            ordered = list(canonical)
            ordered[i], ordered[j] = ordered[j], ordered[i]
            scenarios.append(
                ScenarioSpec(
                    id=slugify_scenario_id("trPr", "swap", canonical[i], canonical[j]),
                    parent_local="trPr",
                    input_children=tuple(ordered),
                    description=f"Swap '{canonical[i]}' and '{canonical[j]}' in CT_TrPr",
                )
            )
    return scenarios


def trpr_full_reverse(canonical: tuple[str, ...]) -> ScenarioSpec:
    """The canonical order, fully reversed. Maximum disturbance — should
    trigger any ordering enforcement Word has."""
    return ScenarioSpec(
        id=slugify_scenario_id("trPr", "reverse"),
        parent_local="trPr",
        input_children=tuple(reversed(canonical)),
        description="CT_TrPr children fully reversed",
    )


# Canonical CT_TrPr child order per ECMA-376 §17.4.79, restricted to the
# core 12 (no w14:conflictIns/Del track-change extensions). Each child
# receives a minimum-valid attribute set from TRPR_CHILD_ATTRS during
# materialization so Word opens the file cleanly — that isolates
# ordering as the only variable.
TRPR_CANONICAL_CORE: tuple[str, ...] = (
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "trHeight",
    "hidden",
    "cantSplit",
    "tblHeader",
    "tblCellSpacing",
    "jc",
)
