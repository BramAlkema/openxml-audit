"""XML-fragment diff helpers for the Word roundtrip oracle.

Pure logic — no Word, no osascript. Lives here (rather than under
`tools/oracle/scenarios/`) because both the engine smoke tests and the
spec-010 scenario driver consume it.

Spec: `specs/011-word-roundtrip-oracle.md`.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass

from lxml import etree


def _local_name(tag: str) -> str:
    if isinstance(tag, str) and tag.startswith("{"):
        return tag.split("}", 1)[1]
    return str(tag)


def extract_first(docx_path: str, parent_local: str, namespace: str) -> list[str] | None:
    """Open a DOCX and return the local-name children of the first matching
    parent element under `word/document.xml`. Returns None if the parent
    is not present.
    """
    with zipfile.ZipFile(docx_path) as zf:
        try:
            doc = zf.read("word/document.xml")
        except KeyError:
            return None
    root = etree.fromstring(doc)
    parent_clark = f"{{{namespace}}}{parent_local}"
    for elem in root.iter(parent_clark):
        return [
            _local_name(child.tag) for child in elem
            if isinstance(child.tag, str)
        ]
    return None


@dataclass
class FragmentDiff:
    """Outcome of comparing one input/post-Word property-element fragment."""

    parent_local: str
    input_children: list[str]
    output_children: list[str]
    verdict: str  # "preserved" | "repaired" | "missing"
    summary: str


def diff_property_fragment(
    parent_local: str,
    input_children: list[str] | None,
    output_children: list[str] | None,
) -> FragmentDiff:
    """Classify the change between input and output children of one
    property-element instance.

    `preserved` — sequences are identical
    `repaired` — sequences differ in any way
    `missing` — the parent element doesn't exist in one side or the other
    """
    if input_children is None or output_children is None:
        return FragmentDiff(
            parent_local=parent_local,
            input_children=input_children or [],
            output_children=output_children or [],
            verdict="missing",
            summary=(
                f"Parent {parent_local!r} not present in "
                f"{'input' if input_children is None else 'output'}"
            ),
        )

    if input_children == output_children:
        return FragmentDiff(
            parent_local=parent_local,
            input_children=input_children,
            output_children=output_children,
            verdict="preserved",
            summary=f"{parent_local}: children identical "
            f"({', '.join(input_children) or '(empty)'})",
        )

    return FragmentDiff(
        parent_local=parent_local,
        input_children=input_children,
        output_children=output_children,
        verdict="repaired",
        summary=(
            f"{parent_local}: input=[{', '.join(input_children)}] "
            f"output=[{', '.join(output_children)}]"
        ),
    )


def slugify_scenario_id(parent_local: str, kind: str, *parts: str) -> str:
    """Build a stable, filesystem-safe ID for an oracle scenario.

    Examples:
        slugify_scenario_id("trPr", "baseline")
            -> "trPr-baseline"
        slugify_scenario_id("trPr", "swap", "cantSplit", "tblHeader")
            -> "trPr-swap-cantSplit-tblHeader"
    """
    pieces = [parent_local, kind, *parts]
    cleaned = []
    for piece in pieces:
        if not piece:
            continue
        # Drop characters that aren't safe in filenames or JSON keys
        safe = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in piece)
        cleaned.append(safe)
    return "-".join(cleaned)


def matches_repair_pattern(text: str, patterns: tuple[str, ...]) -> bool:
    """Case-insensitive substring match for repair-dialog detection.

    Extracted into pure logic so the engine's pattern list can evolve
    without retesting end-to-end Word behavior.
    """
    lowered = text.lower()
    return any(p.lower() in lowered for p in patterns)
