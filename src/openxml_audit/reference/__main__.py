"""CLI for the canonical reference documents (Spec 034).

Usage:
    python -m openxml_audit.reference build --format all \
        --minimum-tier loadable --out build/reference
    python -m openxml_audit.reference status [--json]
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Iterable
from pathlib import Path
from typing import Any

from openxml_audit.evidence import EvidenceTier
from openxml_audit.reference.documents import (
    ReferenceBuildError,
    build_reference_document,
)
from openxml_audit.reference.emitters import has_emitter
from openxml_audit.reference.ledger import (
    FORMATS,
    TIER_ORDER,
    collect_ledger,
    qualifies_at,
)

__all__ = ["main"]


def _parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="python -m openxml_audit.reference", description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    build = subparsers.add_parser(
        "build", help="Build reference documents from the capability ledger."
    )
    build.add_argument(
        "--format",
        choices=[*FORMATS, "all"],
        default="all",
        help="Format to build (default: all).",
    )
    build.add_argument(
        "--minimum-tier",
        choices=[tier.value for tier in TIER_ORDER],
        default=EvidenceTier.ROUNDTRIP_PRESERVED.value,
        help="Minimum evidence tier for inclusion (default: roundtrip-preserved).",
    )
    build.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Output directory for documents and manifests.",
    )

    status = subparsers.add_parser(
        "status", help="Report ledger coverage and emitter gaps per format."
    )
    status.add_argument("--json", action="store_true", help="Emit JSON.")

    return parser.parse_args(list(argv) if argv is not None else None)


def _run_build(args: argparse.Namespace) -> int:
    formats = FORMATS if args.format == "all" else (args.format,)
    minimum_tier = EvidenceTier(args.minimum_tier)
    for fmt in formats:
        try:
            result = build_reference_document(fmt, args.out, minimum_tier=minimum_tier)
        except ReferenceBuildError as exc:
            print(f"{fmt}: BUILD FAILED — {exc}", file=sys.stderr)
            return 1
        included = len(result.included_keys)
        print(
            f"{fmt}: {result.document_path} "
            f"({included} feature{'s' if included != 1 else ''}, "
            f"{len(result.excluded)} excluded) — validated"
        )
        for key, reason in result.excluded:
            print(f"  excluded {key}: {reason}")
    return 0


def _status_payload() -> dict[str, Any]:
    payload: dict[str, Any] = {}
    entries = collect_ledger()
    for fmt in FORMATS:
        fmt_entries = [entry for entry in entries if entry.format == fmt]
        tier_counts = {
            tier.value: sum(1 for entry in fmt_entries if tier in entry.finding.evidence_tiers)
            for tier in TIER_ORDER
        }
        gaps = sorted(
            entry.finding.key
            for entry in fmt_entries
            if qualifies_at(entry.finding, EvidenceTier.LOADABLE)
            and not has_emitter(fmt, entry.finding.key)
        )
        payload[fmt] = {
            "findings": len(fmt_entries),
            "registered_tier_counts": tier_counts,
            "qualifying_at_loadable": sum(
                1 for entry in fmt_entries if qualifies_at(entry.finding, EvidenceTier.LOADABLE)
            ),
            "qualifying_at_roundtrip_preserved": sum(
                1
                for entry in fmt_entries
                if qualifies_at(entry.finding, EvidenceTier.ROUNDTRIP_PRESERVED)
            ),
            "emitter_gaps": gaps,
        }
    return payload


def _run_status(args: argparse.Namespace) -> int:
    payload = _status_payload()
    if args.json:
        print(json.dumps(payload, indent=2, sort_keys=True))
        return 0
    for fmt, stats in payload.items():
        print(f"{fmt}:")
        print(f"  findings: {stats['findings']}")
        print(f"  qualifying at loadable: {stats['qualifying_at_loadable']}")
        print(f"  qualifying at roundtrip-preserved: {stats['qualifying_at_roundtrip_preserved']}")
        if stats["emitter_gaps"]:
            print("  emitter gaps:")
            for key in stats["emitter_gaps"]:
                print(f"    - {key}")
        else:
            print("  emitter gaps: none")
    return 0


def main(argv: Iterable[str] | None = None) -> int:
    args = _parse_args(argv)
    if args.command == "build":
        return _run_build(args)
    return _run_status(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
