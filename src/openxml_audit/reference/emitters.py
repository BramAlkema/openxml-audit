"""Bindings from capability keys to reference-document fragments.

A capability finding earns a place in a reference document only when it
has an emitter here. Findings without emitters are reported as gaps by
the builder and the `status` CLI — never silently dropped.

PPTX features bind to committed oracle scaffold slides
(`data/pptx_oracle/scaffolds/`), which already embed the authored
evidence-bearing XML (timing probes, visual controls). DOCX/XLSX
features bind to callables producing body blocks / sheet rows; both
registries are empty until their capability registries gain findings
(Spec 034 Phase 4).
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from lxml import etree

__all__ = [
    "DOCX_BODY_EMITTERS",
    "PPTX_SLIDE_SOURCES",
    "XLSX_ROW_EMITTERS",
    "PptxSlideSource",
    "has_emitter",
]


@dataclass(frozen=True, slots=True)
class PptxSlideSource:
    """One committed scaffold slide carrying a feature's authored XML."""

    scaffold: str
    slide_number: int


# Slide bindings into the committed `timing_oracle` scaffold, whose
# slides embed the authored probes these findings were registered from.
PPTX_SLIDE_SOURCES: dict[str, tuple[PptxSlideSource, ...]] = {
    "pptx.timing.end-condition.time-offset": (PptxSlideSource("timing_oracle", 2),),
    "pptx.timing.end-condition.click": (PptxSlideSource("timing_oracle", 3),),
    "pptx.timing.repeat-duration": (PptxSlideSource("timing_oracle", 4),),
    "pptx.timing.restart": (
        PptxSlideSource("timing_oracle", 5),
        PptxSlideSource("timing_oracle", 6),
    ),
    # pptx.anim.effect.entr.fade / .wipe: no committed scaffold slide
    # exercises the plain entrance structure those findings describe;
    # they surface as emitter gaps until dedicated slides are authored.
}

# DOCX: capability key -> callable returning <w:body> block elements.
DOCX_BODY_EMITTERS: dict[str, Callable[[], list[etree._Element]]] = {}

# XLSX: capability key -> callable receiving a shared-string interner
# (text -> SST index) and returning <row> elements. Cells must use
# t="s" against the interner — inline strings trigger Excel's silent
# shared-strings rewrite on save (Spec 029 canonical-form finding).
XLSX_ROW_EMITTERS: dict[str, Callable[[Callable[[str], int]], list[etree._Element]]] = {}

_EMITTER_KEYS: dict[str, frozenset[str]] = {
    "pptx": frozenset(PPTX_SLIDE_SOURCES),
    "docx": frozenset(DOCX_BODY_EMITTERS),
    "xlsx": frozenset(XLSX_ROW_EMITTERS),
}


def has_emitter(fmt: str, key: str) -> bool:
    """True when a capability key has a document emitter for `fmt`."""
    return key in _EMITTER_KEYS.get(fmt, frozenset())
