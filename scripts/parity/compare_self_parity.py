#!/usr/bin/env python3
"""Compare a current self-parity snapshot against the committed baseline.

Self-parity (Spec 026, prep for 0.8.0) keys regression detection on
**our own** validator output instead of the .NET SDK's. The
comparator answers three questions per release:

1. **New families** — `family_key`s present in current but not in
   baseline. Each one is a finding the validator now emits that it
   didn't before. Must be intentional.
2. **Missing families** — `family_key`s in baseline but absent in
   current. Improvements (we no longer emit a false positive) AND
   regressions (we lost an intentional finding) both look the same
   here.
3. **Count drift** — same family appears in both, but with
   different counts. Could be a corpus change OR a validator
   change that affects the per-file detection rate.

Threshold flags follow the same shape as `compare_to_baseline.py`
(SDK comparator) so CI workflows can use the same UX.

Spec 013 OQ1 says the right baseline format is "family-set + per-
family count" — that's what this consumes.
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path


@dataclass
class FamilyDiff:
    family_key: str
    baseline_count: int
    current_count: int
    delta: int
    source_class: str = ""


@dataclass
class CompareReport:
    schema_version: int = 1
    baseline_path: str = ""
    current_path: str = ""
    baseline_total: int = 0
    current_total: int = 0
    new_families: list[FamilyDiff] = field(default_factory=list)
    missing_families: list[FamilyDiff] = field(default_factory=list)
    count_drift: list[FamilyDiff] = field(default_factory=list)
    by_source_class_baseline: dict[str, int] = field(default_factory=dict)
    by_source_class_current: dict[str, int] = field(default_factory=dict)
    gate: dict = field(default_factory=dict)


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def compare(
    baseline: dict,
    current: dict,
) -> CompareReport:
    base_inv = baseline.get("family_inventory", {})
    cur_inv = current.get("family_inventory", {})

    base_keys = set(base_inv)
    cur_keys = set(cur_inv)

    new_families = sorted(cur_keys - base_keys)
    missing_families = sorted(base_keys - cur_keys)
    shared = base_keys & cur_keys

    new_diffs = [
        FamilyDiff(
            family_key=k,
            baseline_count=0,
            current_count=cur_inv[k]["count"],
            delta=cur_inv[k]["count"],
            source_class=cur_inv[k].get("source_class", ""),
        )
        for k in new_families
    ]
    missing_diffs = [
        FamilyDiff(
            family_key=k,
            baseline_count=base_inv[k]["count"],
            current_count=0,
            delta=-base_inv[k]["count"],
            source_class=base_inv[k].get("source_class", ""),
        )
        for k in missing_families
    ]
    drift_diffs = [
        FamilyDiff(
            family_key=k,
            baseline_count=base_inv[k]["count"],
            current_count=cur_inv[k]["count"],
            delta=cur_inv[k]["count"] - base_inv[k]["count"],
            source_class=cur_inv[k].get("source_class", base_inv[k].get("source_class", "")),
        )
        for k in sorted(shared)
        if base_inv[k]["count"] != cur_inv[k]["count"]
    ]

    return CompareReport(
        baseline_path=str(baseline.get("manifest", "")),
        current_path=str(current.get("manifest", "")),
        baseline_total=baseline.get("total_findings", 0),
        current_total=current.get("total_findings", 0),
        new_families=new_diffs,
        missing_families=missing_diffs,
        count_drift=drift_diffs,
        by_source_class_baseline=baseline.get("by_source_class", {}),
        by_source_class_current=current.get("by_source_class", {}),
    )


def evaluate_gate(
    report: CompareReport,
    *,
    max_new_families: int = 0,
    max_missing_families: int = 0,
    max_count_drift_total: int = 0,
) -> dict:
    """Return a structured gate verdict. The same shape `compare_to_baseline.py`
    emits, scaled to self-parity's three threshold knobs."""
    new_count = len(report.new_families)
    missing_count = len(report.missing_families)
    drift_total = sum(abs(d.delta) for d in report.count_drift)

    failures: list[str] = []
    if new_count > max_new_families:
        failures.append(
            f"new families exceeded threshold: {new_count} > {max_new_families}"
        )
    if missing_count > max_missing_families:
        failures.append(
            f"missing families exceeded threshold: {missing_count} > {max_missing_families}"
        )
    if drift_total > max_count_drift_total:
        failures.append(
            f"per-family count drift exceeded threshold: {drift_total} > {max_count_drift_total}"
        )

    return {
        "ok": not failures,
        "failures": failures,
        "thresholds": {
            "max_new_families": max_new_families,
            "max_missing_families": max_missing_families,
            "max_count_drift_total": max_count_drift_total,
        },
        "counts": {
            "new_families": new_count,
            "missing_families": missing_count,
            "count_drift_total": drift_total,
        },
    }


def render_summary_md(report: CompareReport) -> str:
    lines: list[str] = []
    lines.append("# Self-parity comparison")
    lines.append("")
    lines.append(f"- baseline total findings: {report.baseline_total}")
    lines.append(f"- current total findings:  {report.current_total}")
    lines.append(f"- new families:     {len(report.new_families)}")
    lines.append(f"- missing families: {len(report.missing_families)}")
    lines.append(f"- count drift:      {len(report.count_drift)} families "
                 f"(|Δ| total = {sum(abs(d.delta) for d in report.count_drift)})")
    lines.append("")
    if report.gate:
        lines.append(f"**Gate verdict:** {'PASS' if report.gate.get('ok') else 'FAIL'}")
        for fail in report.gate.get("failures", []):
            lines.append(f"  - {fail}")
        lines.append("")
    if report.new_families:
        lines.append("## New families (not in baseline)")
        for d in report.new_families[:20]:
            lines.append(f"- ({d.source_class or '?'}) +{d.current_count}  `{d.family_key}`")
        if len(report.new_families) > 20:
            lines.append(f"... and {len(report.new_families) - 20} more")
        lines.append("")
    if report.missing_families:
        lines.append("## Missing families (in baseline but not in current)")
        for d in report.missing_families[:20]:
            lines.append(f"- ({d.source_class or '?'}) -{d.baseline_count}  `{d.family_key}`")
        if len(report.missing_families) > 20:
            lines.append(f"... and {len(report.missing_families) - 20} more")
        lines.append("")
    if report.count_drift:
        lines.append("## Count drift (shared families with different counts)")
        for d in report.count_drift[:20]:
            sign = "+" if d.delta > 0 else ""
            lines.append(
                f"- ({d.source_class or '?'}) {sign}{d.delta}  "
                f"(was {d.baseline_count} → now {d.current_count})  `{d.family_key}`"
            )
        if len(report.count_drift) > 20:
            lines.append(f"... and {len(report.count_drift) - 20} more")
    return "\n".join(lines) + "\n"


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--baseline", required=True, type=Path,
                   help="path to the committed self-parity baseline JSON")
    p.add_argument("--current", required=True, type=Path,
                   help="path to the freshly-generated current snapshot JSON")
    p.add_argument("--output", type=Path, default=None,
                   help="optional path to write the comparison JSON")
    p.add_argument("--summary", type=Path, default=None,
                   help="optional path to write a markdown summary")
    p.add_argument("--max-new-families", type=int, default=0,
                   help="hard fail if new families exceed this (default 0)")
    p.add_argument("--max-missing-families", type=int, default=0,
                   help="hard fail if missing families exceed this (default 0)")
    p.add_argument("--max-count-drift-total", type=int, default=0,
                   help="hard fail if total |Δ| of count drift exceeds this (default 0)")
    args = p.parse_args()

    if not args.baseline.exists():
        print(f"baseline not found: {args.baseline}", file=sys.stderr)
        return 2
    if not args.current.exists():
        print(f"current not found: {args.current}", file=sys.stderr)
        return 2

    baseline = _load(args.baseline)
    current = _load(args.current)
    report = compare(baseline, current)
    report.gate = evaluate_gate(
        report,
        max_new_families=args.max_new_families,
        max_missing_families=args.max_missing_families,
        max_count_drift_total=args.max_count_drift_total,
    )

    payload = {
        "schema_version": report.schema_version,
        "baseline_path": report.baseline_path,
        "current_path": report.current_path,
        "baseline_total": report.baseline_total,
        "current_total": report.current_total,
        "by_source_class_baseline": report.by_source_class_baseline,
        "by_source_class_current": report.by_source_class_current,
        "new_families": [asdict(d) for d in report.new_families],
        "missing_families": [asdict(d) for d in report.missing_families],
        "count_drift": [asdict(d) for d in report.count_drift],
        "gate": report.gate,
    }

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        print(f"wrote comparison: {args.output}", file=sys.stderr)
    else:
        print(json.dumps(payload, indent=2))

    if args.summary:
        args.summary.parent.mkdir(parents=True, exist_ok=True)
        args.summary.write_text(render_summary_md(report), encoding="utf-8")
        print(f"wrote summary: {args.summary}", file=sys.stderr)

    if not report.gate["ok"]:
        print("Gate failed:", file=sys.stderr)
        for fail in report.gate["failures"]:
            print(f"  - {fail}", file=sys.stderr)
        return 1
    print("Gate passed.", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
