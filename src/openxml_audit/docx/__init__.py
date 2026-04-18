"""DOCX-specific oracle layer.

Mirrors the PPTX layout: capability registry wired to the shared evidence
ladder, calibration-emitter starters for minimal probes, and osa-based
Word automation for tier-escalation via re-save.
"""

from __future__ import annotations

from typing import Any


def __getattr__(name: str) -> Any:
    if name in {
        "check_capability",
        "get_capability_finding",
        "list_capability_findings",
    }:
        from openxml_audit.docx import capabilities as _capabilities

        return getattr(_capabilities, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    "check_capability",
    "get_capability_finding",
    "list_capability_findings",
]
