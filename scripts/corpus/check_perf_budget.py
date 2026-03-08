#!/usr/bin/env python3
"""Validate parity snapshot runtime against a configurable performance budget."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

DEFAULT_REPORT = Path("reports/parity_current.json")
DEFAULT_BUDGET = Path("data/corpus/parity_baseline/v3.4.1/perf_budget.json")
DEFAULT_OUTPUT = Path("reports/parity_perf.json")


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _as_int(payload: dict[str, Any], key: str, default: int = 0) -> int:
    value = payload.get(key, default)
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str) and value.isdigit():
        return int(value)
    return default


def _as_float(payload: dict[str, Any], key: str, default: float = 0.0) -> float:
    value = payload.get(key, default)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            return default
    return default


def _build_summary(result: dict[str, Any], failures: list[str]) -> str:
    actual = result["actual"]
    limits = result["limits"]
    lines = [
        "# Parity Performance Summary",
        f"- Duration: {actual['duration_seconds']:.4f}s "
        f"(limit {limits['max_duration_seconds']:.4f}s)",
        f"- Seconds per check: {actual['seconds_per_check']:.6f}s "
        f"(limit {limits['max_seconds_per_check']:.6f}s)",
        f"- Seconds per validation run: {actual['seconds_per_validation']:.6f}s "
        f"(limit {limits['max_seconds_per_validation']:.6f}s)",
        f"- Checks total: {actual['checks_total']} "
        f"(minimum {limits['min_checks_total']})",
        f"- Validation runs: {actual['validation_runs']}",
    ]
    if failures:
        lines.append("")
        lines.append("## Budget result: FAILED")
        for failure in failures:
            lines.append(f"- {failure}")
    else:
        lines.append("")
        lines.append("## Budget result: PASSED")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Check parity snapshot performance budget.")
    parser.add_argument(
        "--report",
        type=Path,
        default=DEFAULT_REPORT,
        help=f"Parity snapshot report path (default: {DEFAULT_REPORT})",
    )
    parser.add_argument(
        "--budget",
        type=Path,
        default=DEFAULT_BUDGET,
        help=f"Performance budget JSON path (default: {DEFAULT_BUDGET})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Performance check output JSON path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=None,
        help="Optional markdown summary output path.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Evaluate and print result without writing output files.",
    )
    args = parser.parse_args()

    report_path = args.report.resolve()
    budget_path = args.budget.resolve()
    if not report_path.exists():
        print(f"Report not found: {report_path}")
        return 2
    if not budget_path.exists():
        print(f"Budget file not found: {budget_path}")
        return 2

    report = _load_json(report_path)
    budget = _load_json(budget_path)
    limits = budget.get("limits")
    if not isinstance(limits, dict):
        print(f"Budget file missing 'limits' object: {budget_path}")
        return 2

    duration_seconds = _as_float(report, "duration_seconds", default=-1.0)
    checks_total = _as_int(report, "checks_total", default=0)
    validation_runs = _as_int(report, "validation_runs", default=0)

    max_duration_seconds = _as_float(limits, "max_duration_seconds", default=0.0)
    max_seconds_per_check = _as_float(limits, "max_seconds_per_check", default=0.0)
    max_seconds_per_validation = _as_float(limits, "max_seconds_per_validation", default=0.0)
    min_checks_total = _as_int(limits, "min_checks_total", default=0)

    seconds_per_check = duration_seconds / checks_total if checks_total > 0 else float("inf")
    seconds_per_validation = (
        duration_seconds / validation_runs if validation_runs > 0 else float("inf")
    )

    failures: list[str] = []
    if duration_seconds < 0:
        failures.append("Snapshot report missing valid duration_seconds.")
    if checks_total < min_checks_total:
        failures.append(
            f"checks_total below minimum: {checks_total} < {min_checks_total}."
        )
    if max_duration_seconds > 0 and duration_seconds > max_duration_seconds:
        failures.append(
            "duration_seconds exceeded limit: "
            f"{duration_seconds:.4f}s > {max_duration_seconds:.4f}s."
        )
    if max_seconds_per_check > 0 and seconds_per_check > max_seconds_per_check:
        failures.append(
            "seconds_per_check exceeded limit: "
            f"{seconds_per_check:.6f}s > {max_seconds_per_check:.6f}s."
        )
    if max_seconds_per_validation > 0 and seconds_per_validation > max_seconds_per_validation:
        failures.append(
            "seconds_per_validation exceeded limit: "
            f"{seconds_per_validation:.6f}s > {max_seconds_per_validation:.6f}s."
        )

    result = {
        "budget_path": str(budget_path),
        "report_path": str(report_path),
        "actual": {
            "duration_seconds": duration_seconds,
            "checks_total": checks_total,
            "validation_runs": validation_runs,
            "seconds_per_check": seconds_per_check,
            "seconds_per_validation": seconds_per_validation,
        },
        "limits": {
            "max_duration_seconds": max_duration_seconds,
            "max_seconds_per_check": max_seconds_per_check,
            "max_seconds_per_validation": max_seconds_per_validation,
            "min_checks_total": min_checks_total,
        },
        "passed": not failures,
        "failures": failures,
    }

    print(f"Duration: {duration_seconds:.4f}s (limit {max_duration_seconds:.4f}s)")
    print(f"Seconds/check: {seconds_per_check:.6f}s (limit {max_seconds_per_check:.6f}s)")
    print(
        "Seconds/validation run: "
        f"{seconds_per_validation:.6f}s (limit {max_seconds_per_validation:.6f}s)"
    )
    print(f"Checks total: {checks_total} (minimum {min_checks_total})")

    summary_content = _build_summary(result=result, failures=failures)
    if args.summary is not None and not args.dry_run:
        summary_path = args.summary.resolve()
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        summary_path.write_text(summary_content, encoding="utf-8")
        print(f"Wrote summary: {summary_path}")

    if not args.dry_run:
        output_path = args.output.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(result, indent=2), encoding="utf-8")
        print(f"Wrote perf report: {output_path}")

    if failures:
        print("Performance budget failed:")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("Performance budget passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
