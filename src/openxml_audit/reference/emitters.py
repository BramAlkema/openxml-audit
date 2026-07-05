"""Bindings from capability keys to reference-document fragments.

A capability finding earns a place in a reference document only when it
has an emitter here. Findings without emitters are reported as gaps by
the builder and the `status` CLI — never silently dropped.

PPTX features bind either to committed oracle scaffold slides
(`data/pptx_oracle/scaffolds/`), which embed the authored
evidence-bearing XML (timing probes, visual controls), or to slides
generated at build time from the same fragment builders the findings
were registered from (`reference/pptx_slides.py`). DOCX/XLSX features
bind to callables producing body blocks / sheet rows; both registries
are empty until their capability registries gain findings
(Spec 034 Phase 4).
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from lxml import etree

from openxml_audit.reference import pptx_slides

__all__ = [
    "DOCX_BODY_EMITTERS",
    "PPTX_SLIDE_SOURCES",
    "XLSX_ROW_EMITTERS",
    "PptxGeneratedSlide",
    "PptxSlideSource",
    "PptxSource",
    "has_emitter",
]


@dataclass(frozen=True, slots=True)
class PptxSlideSource:
    """One committed scaffold slide carrying a feature's authored XML."""

    scaffold: str
    slide_number: int


@dataclass(frozen=True, slots=True)
class PptxGeneratedSlide:
    """A slide authored at build time from the finding's registered structure.

    The builder returns complete slide XML; the timing fragments come
    from the same oracle-deck builders the finding was registered from.
    """

    builder: Callable[[], bytes]


PptxSource = PptxSlideSource | PptxGeneratedSlide

# Timing findings bind into the committed `timing_oracle` scaffold,
# whose slides embed the authored probes they were registered from.
# Entrance-effect findings bind to generated slides carrying the same
# fragment-builder output the April 2026 slideshow probes used.
PPTX_SLIDE_SOURCES: dict[str, tuple[PptxSource, ...]] = {
    "pptx.anim.effect.entr.fade": (PptxGeneratedSlide(pptx_slides.build_entrance_fade_slide),),
    "pptx.anim.effect.entr.wipe": (PptxGeneratedSlide(pptx_slides.build_entrance_wipe_slide),),
    "pptx.timing.end-condition.time-offset": (PptxSlideSource("timing_oracle", 2),),
    "pptx.timing.end-condition.click": (PptxSlideSource("timing_oracle", 3),),
    "pptx.timing.repeat-duration": (PptxSlideSource("timing_oracle", 4),),
    "pptx.timing.restart": (
        PptxSlideSource("timing_oracle", 5),
        PptxSlideSource("timing_oracle", 6),
    ),
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
