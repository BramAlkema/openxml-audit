"""Curated Excel capability findings for lightweight 'will this work?' checks.

Empty registry at kickoff — findings are added as XLSX oracle cases land
with tier evidence. Follows the PPTX pattern from
`src/openxml_audit/pptx/capabilities.py`.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Iterable
from typing import Any

from openxml_audit.evidence import CapabilityFinding, EvidenceTier

__all__ = [
    "check_capability",
    "get_capability_finding",
    "list_capability_findings",
    "main",
]


_FINDINGS: tuple[CapabilityFinding, ...] = ()

_FINDINGS_BY_KEY = {finding.key: finding for finding in _FINDINGS}
for _finding in _FINDINGS:
    for _alias in _finding.aliases:
        _FINDINGS_BY_KEY.setdefault(_alias, _finding)


def get_capability_finding(key: str) -> CapabilityFinding | None:
    """Return one finding by key or alias."""
    return _FINDINGS_BY_KEY.get(key)


def list_capability_findings(prefix: str | None = None) -> list[CapabilityFinding]:
    """Return canonical findings, optionally filtered by key prefix."""
    findings = list(_FINDINGS)
    if prefix:
        findings = [finding for finding in findings if finding.key.startswith(prefix)]
    return sorted(findings, key=lambda finding: finding.key)


def check_capability(
    key: str,
    *,
    minimum_tier: EvidenceTier | None = None,
) -> dict[str, Any]:
    """Return a structured capability report for one finding."""
    finding = get_capability_finding(key)
    if finding is None:
        return {
            "key": key,
            "known": False,
            "meets_minimum_tier": False if minimum_tier else None,
        }

    return {
        "key": finding.key,
        "known": True,
        "meets_minimum_tier": (finding.meets(minimum_tier) if minimum_tier is not None else None),
        "minimum_tier": minimum_tier.value if minimum_tier is not None else None,
        "finding": finding.to_dict(),
    }


def _parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("key", nargs="?", help="Capability key or alias to inspect.")
    parser.add_argument(
        "--minimum-tier",
        choices=[tier.value for tier in EvidenceTier],
        help="Require at least this evidence tier.",
    )
    parser.add_argument(
        "--prefix",
        help="List findings whose keys start with this prefix.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit JSON instead of text.",
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: Iterable[str] | None = None) -> int:
    args = _parse_args(argv)
    minimum_tier = EvidenceTier(args.minimum_tier) if args.minimum_tier is not None else None

    if args.key is None:
        payload: dict[str, Any] = {
            "findings": [finding.to_dict() for finding in list_capability_findings(args.prefix)]
        }
        if args.json:
            print(json.dumps(payload, indent=2, sort_keys=True))
        else:
            for finding in payload["findings"]:
                print(finding["key"])
        return 0

    payload = check_capability(args.key, minimum_tier=minimum_tier)
    if args.json:
        print(json.dumps(payload, indent=2, sort_keys=True))
    else:
        if payload.get("known", False):
            finding = payload["finding"]
            print(finding["key"])
        else:
            print(f"{payload['key']}: unknown")
    return 0 if payload.get("known", False) else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
