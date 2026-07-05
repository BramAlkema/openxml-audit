"""Canonical reference documents generated from the capability ledger.

ADR-001 names the payoff of the corpus work: one `.pptx`, one `.docx`,
one `.xlsx` aggregating every feature proven at a given evidence tier.
This package is the assembly step — it reads the per-format capability
registries, selects qualifying findings, and emits one reference
document per format plus a provenance manifest (Spec 034).

The builder never asserts evidence: tier claims are read from the
registries and reproduced verbatim in the manifest. Reference documents
are derived artifacts, byte-reproducible from committed sources.
"""

from __future__ import annotations

from openxml_audit.reference.documents import (
    ReferenceBuildError,
    ReferenceBuildResult,
    build_reference_document,
)
from openxml_audit.reference.ledger import (
    TIER_ORDER,
    LedgerEntry,
    collect_ledger,
    qualifies_at,
    tier_rank,
)

__all__ = [
    "TIER_ORDER",
    "LedgerEntry",
    "ReferenceBuildError",
    "ReferenceBuildResult",
    "build_reference_document",
    "collect_ledger",
    "qualifies_at",
    "tier_rank",
]
