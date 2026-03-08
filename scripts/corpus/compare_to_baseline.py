#!/usr/bin/env python3
"""Compare a parity snapshot to a frozen baseline and enforce drift thresholds."""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

DEFAULT_BASELINE = Path("data/corpus/parity_baseline/v3.4.1/parity_snapshot.json")
DEFAULT_CURRENT = Path("reports/parity_snapshot.json")
DEFAULT_OUTPUT = Path("reports/parity_compare.json")
DEFAULT_WAIVERS = Path("data/corpus/parity_baseline/v3.4.1/waivers.json")

WAIVER_KINDS = {
    "new_mismatch_family",
    "mismatch_growth",
    "match_rate_drop",
    "missing_files",
    "check_total_drop",
    "strict_mode",
}


@dataclass(frozen=True)
class SnapshotStats:
    """Normalized parity snapshot fields used for drift comparison."""

    checks_total: int
    checks_matched: int
    checks_mismatched: int
    checks_missing_files: int
    match_rate_percent: float
    strict: bool | None
    mismatch_families: dict[str, int]


@dataclass(frozen=True)
class Waiver:
    """Active waiver entry used by parity gate."""

    kind: str
    target: str | None
    owner: str
    reason: str
    expires: str


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _coerce_int(payload: dict[str, Any], key: str) -> int:
    value = payload.get(key, 0)
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str) and value.isdigit():
        return int(value)
    return 0


def _coerce_float(payload: dict[str, Any], key: str) -> float:
    value = payload.get(key, 0.0)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            return 0.0
    return 0.0


def _extract_families(payload: dict[str, Any]) -> dict[str, int]:
    output: dict[str, int] = {}
    rows = payload.get("mismatch_families")
    if not isinstance(rows, list):
        return output
    for row in rows:
        if not isinstance(row, dict):
            continue
        description = row.get("description")
        count = row.get("count")
        if not isinstance(description, str):
            continue
        if not isinstance(count, int):
            continue
        output[description] = count
    return output


def _load_waivers(path: Path | None) -> tuple[list[Waiver], list[str]]:
    if path is None:
        return [], []
    if not path.exists():
        return [], []

    payload = _load_json(path)
    rows = payload.get("waivers")
    if not isinstance(rows, list):
        return [], [f"Waiver file has invalid 'waivers' payload: {path}"]

    active: list[Waiver] = []
    warnings: list[str] = []
    today = date.today()
    for idx, row in enumerate(rows):
        if not isinstance(row, dict):
            warnings.append(f"Waiver entry {idx} is not an object.")
            continue
        kind = row.get("kind")
        owner = row.get("owner")
        reason = row.get("reason")
        expires = row.get("expires")
        target = row.get("target")

        if not isinstance(kind, str) or kind not in WAIVER_KINDS:
            warnings.append(f"Waiver entry {idx} has invalid kind: {kind!r}.")
            continue
        if not isinstance(owner, str) or not owner.strip():
            warnings.append(f"Waiver entry {idx} missing owner.")
            continue
        if not isinstance(reason, str) or not reason.strip():
            warnings.append(f"Waiver entry {idx} missing reason.")
            continue
        if not isinstance(expires, str):
            warnings.append(f"Waiver entry {idx} missing expires date.")
            continue
        if target is not None and not isinstance(target, str):
            warnings.append(f"Waiver entry {idx} has non-string target.")
            continue
        if kind == "new_mismatch_family" and (not isinstance(target, str) or not target.strip()):
            warnings.append(f"Waiver entry {idx} kind=new_mismatch_family requires target.")
            continue

        try:
            expires_date = date.fromisoformat(expires)
        except ValueError:
            warnings.append(f"Waiver entry {idx} has invalid expires date: {expires!r}.")
            continue
        if expires_date < today:
            warnings.append(
                f"Waiver expired and ignored: kind={kind} target={target!r} expires={expires}."
            )
            continue

        active.append(
            Waiver(
                kind=kind,
                target=target.strip() if isinstance(target, str) and target.strip() else None,
                owner=owner.strip(),
                reason=reason.strip(),
                expires=expires,
            )
        )

    return active, warnings


def _find_waiver(waivers: list[Waiver], kind: str, target: str | None = None) -> Waiver | None:
    for waiver in waivers:
        if waiver.kind != kind:
            continue
        if waiver.target is None:
            return waiver
        if target is not None and waiver.target == target:
            return waiver
    return None


def _load_snapshot(path: Path) -> SnapshotStats:
    payload = _load_json(path)
    strict = payload.get("strict")
    strict_value = strict if isinstance(strict, bool) else None
    return SnapshotStats(
        checks_total=_coerce_int(payload, "checks_total"),
        checks_matched=_coerce_int(payload, "checks_matched"),
        checks_mismatched=_coerce_int(payload, "checks_mismatched"),
        checks_missing_files=_coerce_int(payload, "checks_missing_files"),
        match_rate_percent=_coerce_float(payload, "match_rate_percent"),
        strict=strict_value,
        mismatch_families=_extract_families(payload),
    )


def _family_deltas(
    baseline: dict[str, int], current: dict[str, int]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    new_rows: list[dict[str, Any]] = []
    regressed_rows: list[dict[str, Any]] = []
    improved_rows: list[dict[str, Any]] = []

    all_keys = sorted(set(baseline) | set(current))
    for key in all_keys:
        before = baseline.get(key, 0)
        after = current.get(key, 0)
        delta = after - before
        row = {"description": key, "baseline": before, "current": after, "delta": delta}
        if before == 0 and after > 0:
            new_rows.append(row)
        if delta > 0:
            regressed_rows.append(row)
        elif delta < 0:
            improved_rows.append(row)

    regressed_rows.sort(key=lambda row: int(row["delta"]), reverse=True)
    improved_rows.sort(key=lambda row: int(row["delta"]))
    new_rows.sort(key=lambda row: int(row["current"]), reverse=True)
    return new_rows, regressed_rows, improved_rows


def _build_markdown(
    comparison: dict[str, Any], failures: list[str], waived_failures: list[str]
) -> str:
    lines: list[str] = ["# Parity Drift Summary"]
    lines.append(
        f"- Checks total: {comparison['baseline']['checks_total']} -> "
        f"{comparison['current']['checks_total']} "
        f"(delta {comparison['deltas']['checks_total']})"
    )
    lines.append(
        f"- Mismatched checks: {comparison['baseline']['checks_mismatched']} -> "
        f"{comparison['current']['checks_mismatched']} "
        f"(delta {comparison['deltas']['checks_mismatched']})"
    )
    lines.append(
        f"- Match rate: {comparison['baseline']['match_rate_percent']}% -> "
        f"{comparison['current']['match_rate_percent']}% "
        f"(delta {comparison['deltas']['match_rate_percent']:+.2f}pp)"
    )
    lines.append(
        f"- Missing files: {comparison['baseline']['checks_missing_files']} -> "
        f"{comparison['current']['checks_missing_files']}"
    )
    lines.append(f"- New mismatch families: {comparison['counts']['new_mismatch_families']}")
    lines.append(f"- Waived new mismatch families: {comparison['counts']['waived_new_families']}")

    if failures:
        lines.append("")
        lines.append("## Gate result: FAILED")
        for failure in failures:
            lines.append(f"- {failure}")
    else:
        lines.append("")
        lines.append("## Gate result: PASSED")

    if waived_failures:
        lines.append("")
        lines.append("## Waived conditions")
        for detail in waived_failures:
            lines.append(f"- {detail}")

    regressions = comparison["families"]["regressed"][:10]
    if regressions:
        lines.append("")
        lines.append("## Top regressed families")
        for row in regressions:
            lines.append(
                f"- {row['delta']}x {row['description']} "
                f"(baseline {row['baseline']}, current {row['current']})"
            )

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare parity snapshot with baseline.")
    parser.add_argument(
        "--baseline",
        type=Path,
        default=DEFAULT_BASELINE,
        help=f"Baseline snapshot path (default: {DEFAULT_BASELINE})",
    )
    parser.add_argument(
        "--current",
        type=Path,
        default=DEFAULT_CURRENT,
        help=f"Current snapshot path (default: {DEFAULT_CURRENT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Comparison output JSON path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=None,
        help="Optional markdown summary output path.",
    )
    parser.add_argument(
        "--waivers",
        type=Path,
        default=DEFAULT_WAIVERS,
        help=(
            "Optional waiver file path "
            f"(default: {DEFAULT_WAIVERS}, ignored if missing)."
        ),
    )
    parser.add_argument(
        "--max-mismatch-growth",
        type=int,
        default=0,
        help="Allowed growth in mismatched checks (default: 0).",
    )
    parser.add_argument(
        "--max-new-families",
        type=int,
        default=0,
        help="Allowed count of new mismatch families (default: 0).",
    )
    parser.add_argument(
        "--max-match-rate-drop",
        type=float,
        default=0.0,
        help="Allowed drop in match rate percentage points (default: 0.0).",
    )
    parser.add_argument(
        "--max-missing-files",
        type=int,
        default=0,
        help="Allowed missing file checks in current snapshot (default: 0).",
    )
    parser.add_argument(
        "--allow-check-total-drop",
        action="store_true",
        help="Allow current checks_total to be less than baseline checks_total.",
    )
    parser.add_argument(
        "--require-same-strict",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Require strict mode value to match baseline.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not write output files.",
    )
    args = parser.parse_args()

    baseline_path = args.baseline.resolve()
    current_path = args.current.resolve()
    if not baseline_path.exists():
        print(f"Baseline snapshot not found: {baseline_path}")
        return 2
    if not current_path.exists():
        print(f"Current snapshot not found: {current_path}")
        return 2

    baseline = _load_snapshot(baseline_path)
    current = _load_snapshot(current_path)
    waivers, waiver_warnings = _load_waivers(args.waivers.resolve() if args.waivers else None)

    new_families, regressed_families, improved_families = _family_deltas(
        baseline.mismatch_families,
        current.mismatch_families,
    )
    waived_new_families: list[dict[str, Any]] = []
    unwaived_new_families: list[dict[str, Any]] = []
    for row in new_families:
        waiver = _find_waiver(waivers, "new_mismatch_family", str(row["description"]))
        if waiver is not None:
            waived_new_families.append(
                {
                    **row,
                    "waiver": {
                        "owner": waiver.owner,
                        "reason": waiver.reason,
                        "expires": waiver.expires,
                    },
                }
            )
        else:
            unwaived_new_families.append(row)
    mismatch_growth = current.checks_mismatched - baseline.checks_mismatched
    checks_total_delta = current.checks_total - baseline.checks_total
    match_rate_delta = round(current.match_rate_percent - baseline.match_rate_percent, 4)
    match_rate_drop = baseline.match_rate_percent - current.match_rate_percent

    comparison: dict[str, Any] = {
        "baseline_path": str(baseline_path),
        "current_path": str(current_path),
        "baseline": {
            "checks_total": baseline.checks_total,
            "checks_matched": baseline.checks_matched,
            "checks_mismatched": baseline.checks_mismatched,
            "checks_missing_files": baseline.checks_missing_files,
            "match_rate_percent": baseline.match_rate_percent,
            "strict": baseline.strict,
        },
        "current": {
            "checks_total": current.checks_total,
            "checks_matched": current.checks_matched,
            "checks_mismatched": current.checks_mismatched,
            "checks_missing_files": current.checks_missing_files,
            "match_rate_percent": current.match_rate_percent,
            "strict": current.strict,
        },
        "deltas": {
            "checks_total": checks_total_delta,
            "checks_mismatched": mismatch_growth,
            "match_rate_percent": match_rate_delta,
        },
        "counts": {
            "new_mismatch_families": len(new_families),
            "waived_new_families": len(waived_new_families),
            "unwaived_new_families": len(unwaived_new_families),
            "regressed_families": len(regressed_families),
            "improved_families": len(improved_families),
        },
        "families": {
            "new": new_families,
            "new_waived": waived_new_families,
            "new_unwaived": unwaived_new_families,
            "regressed": regressed_families,
            "improved": improved_families,
        },
        "waiver_warnings": waiver_warnings,
    }

    failures: list[str] = []
    waived_failures: list[str] = []
    if mismatch_growth > args.max_mismatch_growth:
        waiver = _find_waiver(waivers, "mismatch_growth")
        if waiver is not None:
            waived_failures.append(
                "Mismatch growth waiver applied "
                f"(owner={waiver.owner}, expires={waiver.expires})."
            )
        else:
            failures.append(
                "Mismatch growth exceeded threshold: "
                f"{mismatch_growth} > {args.max_mismatch_growth}."
            )
    if len(unwaived_new_families) > args.max_new_families:
        failures.append(
            "New mismatch-family count exceeded threshold: "
            f"{len(unwaived_new_families)} > {args.max_new_families}."
        )
    if match_rate_drop > args.max_match_rate_drop:
        waiver = _find_waiver(waivers, "match_rate_drop")
        if waiver is not None:
            waived_failures.append(
                "Match-rate drop waiver applied "
                f"(owner={waiver.owner}, expires={waiver.expires})."
            )
        else:
            failures.append(
                "Match-rate drop exceeded threshold: "
                f"{match_rate_drop:.2f}pp > {args.max_match_rate_drop:.2f}pp."
            )
    if current.checks_missing_files > args.max_missing_files:
        waiver = _find_waiver(waivers, "missing_files")
        if waiver is not None:
            waived_failures.append(
                "Missing-files waiver applied "
                f"(owner={waiver.owner}, expires={waiver.expires})."
            )
        else:
            failures.append(
                "Missing-file check count exceeded threshold: "
                f"{current.checks_missing_files} > {args.max_missing_files}."
            )
    if not args.allow_check_total_drop and current.checks_total < baseline.checks_total:
        waiver = _find_waiver(waivers, "check_total_drop")
        if waiver is not None:
            waived_failures.append(
                "Check-total-drop waiver applied "
                f"(owner={waiver.owner}, expires={waiver.expires})."
            )
        else:
            failures.append(
                "Current checks_total is lower than baseline: "
                f"{current.checks_total} < {baseline.checks_total}."
            )
    if args.require_same_strict and baseline.strict != current.strict:
        waiver = _find_waiver(waivers, "strict_mode")
        if waiver is not None:
            waived_failures.append(
                "Strict-mode waiver applied "
                f"(owner={waiver.owner}, expires={waiver.expires})."
            )
        else:
            failures.append(
                f"Strict mode changed: baseline={baseline.strict} current={current.strict}."
            )

    comparison["gate"] = {
        "passed": not failures,
        "failures": failures,
        "thresholds": {
            "max_mismatch_growth": args.max_mismatch_growth,
            "max_new_families": args.max_new_families,
            "max_match_rate_drop": args.max_match_rate_drop,
            "max_missing_files": args.max_missing_files,
            "allow_check_total_drop": args.allow_check_total_drop,
            "require_same_strict": args.require_same_strict,
        },
        "waived_conditions": waived_failures,
    }

    print(f"Checks total delta: {checks_total_delta}")
    print(f"Mismatch growth: {mismatch_growth}")
    print(f"Match rate delta: {match_rate_delta:+.2f}pp")
    print(
        "New mismatch families: "
        f"{len(new_families)} (unwaived {len(unwaived_new_families)}, "
        f"waived {len(waived_new_families)})"
    )
    print(f"Missing files (current): {current.checks_missing_files}")
    if waiver_warnings:
        print(f"Waiver warnings: {len(waiver_warnings)}")
        for warning in waiver_warnings:
            print(f"- {warning}")
    if waived_failures:
        print("Waived conditions:")
        for detail in waived_failures:
            print(f"- {detail}")

    summary_content = _build_markdown(
        comparison=comparison,
        failures=failures,
        waived_failures=waived_failures,
    )
    if args.summary is not None and not args.dry_run:
        summary_path = args.summary.resolve()
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        summary_path.write_text(summary_content, encoding="utf-8")
        print(f"Wrote summary: {summary_path}")

    if not args.dry_run:
        output_path = args.output.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(comparison, indent=2), encoding="utf-8")
        print(f"Wrote comparison: {output_path}")

    if failures:
        print("Gate failed:")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("Gate passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
