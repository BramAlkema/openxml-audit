"""Evidence ladder primitives: tiers of proof and capability findings.

The ladder is the organizing principle of this project: validation (schema
legality), loadability, roundtrip preservation, runtime behavior, and authoring
provenance are tiers of one question — "will this file survive?" — not
separate concerns.

These types are format-neutral. PPTX/DOCX/XLSX-specific finding registries live
under their respective format packages and reference these primitives.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any

__all__ = ["CapabilityFinding", "EvidenceTier"]


class EvidenceTier(str, Enum):
    """Progressively stronger evidence tiers for native Office behavior."""

    SCHEMA_VALID = "schema-valid"
    LOADABLE = "loadable"
    ROUNDTRIP_PRESERVED = "roundtrip-preserved"
    SLIDESHOW_VERIFIED = "slideshow-verified"
    UI_AUTHORED = "ui-authored"


@dataclass(frozen=True, slots=True)
class CapabilityFinding:
    """One empirical finding about a native Office feature."""

    key: str
    summary: str
    evidence_tiers: tuple[EvidenceTier, ...]
    constraints: tuple[str, ...] = ()
    notes: tuple[str, ...] = ()
    calibration_artifacts: tuple[str, ...] = ()
    aliases: tuple[str, ...] = field(default_factory=tuple)

    def meets(self, minimum_tier: EvidenceTier) -> bool:
        return minimum_tier in self.evidence_tiers

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = asdict(self)
        payload["evidence_tiers"] = [tier.value for tier in self.evidence_tiers]
        return payload
