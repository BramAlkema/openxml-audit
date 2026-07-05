"""Cross-format capability ledger for reference-document assembly.

Aggregates the per-format capability registries into one view and fixes
the ladder ranking used to answer "does this finding qualify at tier T
or above?". Registered tier tuples remain sets of explicit claims — the
ledger ranks them for selection but never rewrites them.
"""

from __future__ import annotations

from dataclasses import dataclass

from openxml_audit.docx import capabilities as docx_capabilities
from openxml_audit.evidence import CapabilityFinding, EvidenceTier
from openxml_audit.pptx import capabilities as pptx_capabilities
from openxml_audit.xlsx import capabilities as xlsx_capabilities

__all__ = [
    "FORMATS",
    "TIER_ORDER",
    "LedgerEntry",
    "collect_ledger",
    "qualifies_at",
    "tier_rank",
]

FORMATS = ("docx", "pptx", "xlsx")

# ADR-001 ladder order, weakest to strongest.
TIER_ORDER = (
    EvidenceTier.SCHEMA_VALID,
    EvidenceTier.LOADABLE,
    EvidenceTier.ROUNDTRIP_PRESERVED,
    EvidenceTier.SLIDESHOW_VERIFIED,
    EvidenceTier.UI_AUTHORED,
)

_TIER_RANK = {tier: rank for rank, tier in enumerate(TIER_ORDER)}

_REGISTRIES = {
    "docx": docx_capabilities.list_capability_findings,
    "pptx": pptx_capabilities.list_capability_findings,
    "xlsx": xlsx_capabilities.list_capability_findings,
}


def tier_rank(tier: EvidenceTier) -> int:
    """Return the ladder rank of a tier (0 = schema-valid)."""
    return _TIER_RANK[tier]


def qualifies_at(finding: CapabilityFinding, minimum_tier: EvidenceTier) -> bool:
    """True when any registered tier ranks at or above `minimum_tier`."""
    minimum_rank = tier_rank(minimum_tier)
    return any(tier_rank(tier) >= minimum_rank for tier in finding.evidence_tiers)


@dataclass(frozen=True, slots=True)
class LedgerEntry:
    """One capability finding in cross-format ledger context."""

    format: str
    finding: CapabilityFinding

    @property
    def highest_tier(self) -> EvidenceTier | None:
        if not self.finding.evidence_tiers:
            return None
        return max(self.finding.evidence_tiers, key=tier_rank)


def collect_ledger(formats: tuple[str, ...] = FORMATS) -> list[LedgerEntry]:
    """Return ledger entries for the requested formats, sorted by key."""
    entries: list[LedgerEntry] = []
    for fmt in formats:
        if fmt not in _REGISTRIES:
            raise ValueError(f"Unknown reference format: {fmt!r}")
        entries.extend(LedgerEntry(format=fmt, finding=finding) for finding in _REGISTRIES[fmt]())
    return sorted(entries, key=lambda entry: (entry.format, entry.finding.key))
