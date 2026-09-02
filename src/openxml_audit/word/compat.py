"""Historical WordprocessingML ordering proxy used by research tools.

Spec 010 originally used these XSD-derived child sequences as runtime
Word-compatibility warnings. The Word oracle later preserved all 396 tested
ordering scenarios without a repair dialog or XML rewrite, including issue
#3's exact example. The runtime warning was therefore retired in 0.8.0.

The tables and subsequence helper remain available to the empirical mining
tool so future evidence can be compared with the original hypothesis. They
must not be wired into document validation without new, versioned Word-oracle
evidence that demonstrates an actual repair boundary.
"""

from __future__ import annotations

from dataclasses import dataclass

from openxml_audit.namespaces import WORDPROCESSINGML

W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"


def _w(local: str) -> str:
    return f"{{{WORDPROCESSINGML}}}{local}"


def _w14(local: str) -> str:
    return f"{{{W14_NS}}}{local}"


def _w15(local: str) -> str:
    return f"{{{W15_NS}}}{local}"


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
# `w:CT_TrPr/w:trPr` Children. The constraint shipped under the
# assumption (per issue #3) that Word's repair dialog enforces this
# order. The current oracle baseline
# (tools/oracle/baselines/word_trpr_pairwise.json) covers all 12
# canonical children with minimum-valid attributes and 68 scenarios
# (baseline + 66 pairwise swaps + full reverse): Word for Mac M365
# 16.89.1 preserves every ordering with no repair dialog, including
# the canonical fully reversed and issue #3's exact pattern. The
# constraint as written may flag only false positives on that Word
# build. WARNING severity is intentionally conservative; do not
# promote to ERROR until oracle runs on additional Word builds (esp.
# Windows or earlier macOS releases) demonstrate ordering enforcement.
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


# CT_TblPr canonical child order, derived from SDK schema metadata
# `w:CT_TblPr/w:tblPr` Children. Initially validated against TokenMoulds
# corpus (4,847 observations, 100% pass against the proxy). The oracle
# baseline at `tools/oracle/baselines/word_tblpr_pairwise.json` then
# tested all 93 ordering scenarios (baseline + 91 pairwise swaps + full
# reverse) against Word for Mac M365 16.89.1 and recorded 93 preserved,
# 0 repaired, 0 dialogs. The constraint as written may flag false
# positives on this Word build; WARNING severity remains correct.
_CT_TBL_PR = ChildSequence(
    parent_tag=_w("tblPr"),
    parent_local="tblPr",
    spec_section="ECMA-376 §17.4.60",
    children=(
        _w("tblStyle"),
        _w("tblpPr"),
        _w("tblOverlap"),
        _w("bidiVisual"),
        _w("tblW"),
        _w("jc"),
        _w("tblCellSpacing"),
        _w("tblInd"),
        _w("tblBorders"),
        _w("shd"),
        _w("tblLayout"),
        _w("tblCellMar"),
        _w("tblLook"),
        _w("tblCaption"),
        _w("tblDescription"),
        _w("tblPrChange"),
    ),
)


# CT_TcPr canonical child order, derived from SDK schema metadata
# `w:CT_TcPr/w:tcPr` Children. Initially validated against TokenMoulds
# corpus (28,667 observations, 100% pass against the proxy). The oracle
# baseline at `tools/oracle/baselines/word_tcpr_pairwise.json` then
# tested all 80 ordering scenarios (baseline + 78 pairwise swaps + full
# reverse) against Word for Mac M365 16.89.1 and recorded 80 preserved,
# 0 repaired, 0 dialogs. The constraint as written may flag false
# positives on this Word build; WARNING severity remains correct.
_CT_TC_PR = ChildSequence(
    parent_tag=_w("tcPr"),
    parent_local="tcPr",
    spec_section="ECMA-376 §17.4.70",
    children=(
        _w("cnfStyle"),
        _w("tcW"),
        _w("gridSpan"),
        _w("hMerge"),
        _w("vMerge"),
        _w("tcBorders"),
        _w("shd"),
        _w("noWrap"),
        _w("tcMar"),
        _w("textDirection"),
        _w("tcFitText"),
        _w("vAlign"),
        _w("hideMark"),
        _w("cellIns"),
        _w("cellDel"),
        _w("cellMerge"),
        _w("tcPrChange"),
    ),
)


# CT_SectPr canonical child order, derived from SDK schema metadata
# `w:CT_SectPr/w:sectPr` Children. Initially validated against
# TokenMoulds corpus (40 observations, 100% pass; small sample). The
# oracle baseline at `tools/oracle/baselines/word_sectpr_pairwise.json`
# then tested all 155 ordering scenarios (baseline + 153 pairwise swaps
# + full reverse) against Word for Mac M365 16.89.1 and recorded 155
# preserved, 0 repaired, 0 dialogs. The constraint as written may flag
# false positives on this Word build; WARNING severity remains correct.
_CT_SECT_PR = ChildSequence(
    parent_tag=_w("sectPr"),
    parent_local="sectPr",
    spec_section="ECMA-376 §17.6.18",
    children=(
        _w("headerReference"),
        _w("footerReference"),
        _w("footnotePr"),
        _w("endnotePr"),
        _w("type"),
        _w("pgSz"),
        _w("pgMar"),
        _w("paperSrc"),
        _w("pgBorders"),
        _w("lnNumType"),
        _w("pgNumType"),
        _w("cols"),
        _w("formProt"),
        _w("vAlign"),
        _w("noEndnote"),
        _w("titlePg"),
        _w("textDirection"),
        _w("bidi"),
        _w("rtlGutter"),
        _w("docGrid"),
        _w("printerSettings"),
        _w15("footnoteColumns"),
        _w("sectPrChange"),
    ),
)


CONSTRAINT_TABLE: dict[str, ChildSequence] = {
    _CT_TR_PR.parent_tag: _CT_TR_PR,
    _CT_TBL_PR.parent_tag: _CT_TBL_PR,
    _CT_TC_PR.parent_tag: _CT_TC_PR,
    _CT_SECT_PR.parent_tag: _CT_SECT_PR,
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
