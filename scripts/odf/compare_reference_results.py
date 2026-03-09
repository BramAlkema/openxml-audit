#!/usr/bin/env python3
"""Compare normalized Python ODF findings against reference validator runs."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_INPUT = Path("reports/odf/reference_runs.json")
DEFAULT_OUTPUT = Path("reports/odf/reference_compare.json")
DEFAULT_SUMMARY = Path("reports/odf/reference_compare.md")


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _counter_from_issues(issues: Any) -> Counter[str]:
    counter: Counter[str] = Counter()
    if not isinstance(issues, list):
        return counter
    for issue in issues:
        if not isinstance(issue, dict):
            continue
        key = issue.get("comparison_key")
        if not isinstance(key, str) or not key.strip():
            fallback = issue.get("description")
            if isinstance(fallback, str) and fallback.strip():
                key = fallback
            else:
                continue
        counter[key] += 1
    return counter


def _description_map(issues: Any) -> dict[str, str]:
    mapping: dict[str, str] = {}
    if not isinstance(issues, list):
        return mapping
    for issue in issues:
        if not isinstance(issue, dict):
            continue
        key = issue.get("comparison_key")
        if not isinstance(key, str) or not key:
            continue
        description = issue.get("description")
        if isinstance(description, str) and description and key not in mapping:
            mapping[key] = description
    return mapping


def _sorted_family_rows(
    counter: Counter[str], descriptions: dict[str, str]
) -> list[dict[str, Any]]:
    rows = [
        {
            "comparison_key": key,
            "description": descriptions.get(key, key),
            "count": count,
        }
        for key, count in counter.items()
    ]
    rows.sort(key=lambda row: int(row["count"]), reverse=True)
    return rows


def _collect_tool_names(report: dict[str, Any]) -> list[str]:
    names: set[str] = set()
    runners = report.get("runners")
    if isinstance(runners, dict):
        for key in runners:
            if key != "python":
                names.add(str(key))
    samples = report.get("samples")
    if isinstance(samples, list):
        for sample in samples:
            if not isinstance(sample, dict):
                continue
            runs = sample.get("runs")
            if not isinstance(runs, dict):
                continue
            for key in runs:
                if key != "python":
                    names.add(str(key))
    if not names:
        return ["odf_toolkit", "opf"]
    return sorted(names)


def _compare_tool(report: dict[str, Any], tool_name: str) -> dict[str, Any]:
    samples = report.get("samples")
    if not isinstance(samples, list):
        samples = []

    status_counts: Counter[str] = Counter()
    only_python_families: Counter[str] = Counter()
    only_reference_families: Counter[str] = Counter()
    description_lookup: dict[str, str] = {}

    samples_compared = 0
    samples_skipped = 0
    python_total = 0
    reference_total = 0
    matched_total = 0
    only_python_total = 0
    only_reference_total = 0
    sample_rows: list[dict[str, Any]] = []

    for sample in samples:
        if not isinstance(sample, dict):
            continue
        sample_id = str(sample.get("id", "unknown"))
        runs = sample.get("runs")
        if not isinstance(runs, dict):
            samples_skipped += 1
            sample_rows.append(
                {
                    "id": sample_id,
                    "status": "skipped",
                    "reason": "missing runs payload",
                }
            )
            continue

        python_run = runs.get("python")
        reference_run = runs.get(tool_name)
        if not isinstance(python_run, dict) or not isinstance(reference_run, dict):
            samples_skipped += 1
            sample_rows.append(
                {
                    "id": sample_id,
                    "status": "skipped",
                    "reason": f"missing python/{tool_name} run payload",
                }
            )
            continue

        reference_status = str(reference_run.get("status", "unknown"))
        status_counts[reference_status] += 1
        if str(python_run.get("status")) != "ok" or reference_status != "ok":
            samples_skipped += 1
            sample_rows.append(
                {
                    "id": sample_id,
                    "status": "skipped",
                    "reason": (
                        f"python status={python_run.get('status')} "
                        f"{tool_name} status={reference_status}"
                    ),
                }
            )
            continue

        python_counter = _counter_from_issues(python_run.get("issues"))
        reference_counter = _counter_from_issues(reference_run.get("issues"))
        python_descriptions = _description_map(python_run.get("issues"))
        reference_descriptions = _description_map(reference_run.get("issues"))

        description_lookup.update(reference_descriptions)
        description_lookup.update(python_descriptions)

        shared = python_counter & reference_counter
        only_python = python_counter - reference_counter
        only_reference = reference_counter - python_counter

        samples_compared += 1
        python_issues = sum(python_counter.values())
        reference_issues = sum(reference_counter.values())
        matched = sum(shared.values())
        only_python_count = sum(only_python.values())
        only_reference_count = sum(only_reference.values())

        python_total += python_issues
        reference_total += reference_issues
        matched_total += matched
        only_python_total += only_python_count
        only_reference_total += only_reference_count

        only_python_families.update(only_python)
        only_reference_families.update(only_reference)

        sample_rows.append(
            {
                "id": sample_id,
                "status": "compared",
                "python_issue_count": python_issues,
                "reference_issue_count": reference_issues,
                "matched": matched,
                "only_python": only_python_count,
                "only_reference": only_reference_count,
            }
        )

    return {
        "tool": tool_name,
        "samples_total": len(samples),
        "samples_compared": samples_compared,
        "samples_skipped": samples_skipped,
        "reference_status_counts": dict(status_counts),
        "issue_totals": {
            "python": python_total,
            "reference": reference_total,
            "matched": matched_total,
            "only_python": only_python_total,
            "only_reference": only_reference_total,
        },
        "mismatch_families": {
            "only_python": _sorted_family_rows(only_python_families, description_lookup),
            "only_reference": _sorted_family_rows(only_reference_families, description_lookup),
        },
        "samples": sample_rows,
    }


def _build_markdown(report: dict[str, Any]) -> str:
    lines = ["# ODF Reference Comparison Summary", ""]
    lines.append(f"- Generated at: {report['generated_at']}")
    lines.append(f"- Input report: {report['input_report']}")
    lines.append(f"- Tools compared: {', '.join(report['tool_order'])}")
    lines.append("")

    tools = report.get("tools", {})
    for tool_name in report["tool_order"]:
        tool = tools.get(tool_name, {})
        totals = tool.get("issue_totals", {})
        lines.append(f"## {tool_name}")
        lines.append(f"- Samples total: {tool.get('samples_total', 0)}")
        lines.append(f"- Samples compared: {tool.get('samples_compared', 0)}")
        lines.append(f"- Samples skipped: {tool.get('samples_skipped', 0)}")
        lines.append(f"- Reference status counts: {tool.get('reference_status_counts', {})}")
        lines.append(
            "- Issue totals: "
            f"python={totals.get('python', 0)}, "
            f"reference={totals.get('reference', 0)}, "
            f"matched={totals.get('matched', 0)}, "
            f"only_python={totals.get('only_python', 0)}, "
            f"only_reference={totals.get('only_reference', 0)}"
        )

        only_python_rows = tool.get("mismatch_families", {}).get("only_python", [])[:10]
        only_reference_rows = tool.get("mismatch_families", {}).get("only_reference", [])[:10]

        if only_python_rows:
            lines.append("")
            lines.append("### Top only-python families")
            for row in only_python_rows:
                lines.append(f"- {row['count']}x {row['description']}")

        if only_reference_rows:
            lines.append("")
            lines.append("### Top only-reference families")
            for row in only_reference_rows:
                lines.append(f"- {row['count']}x {row['description']}")

        if not only_python_rows and not only_reference_rows:
            lines.append("- No mismatch families recorded.")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare ODF reference run results.")
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help=f"Input run report from run_reference_validators.py (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output JSON comparison report (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=DEFAULT_SUMMARY,
        help=f"Output markdown summary report (default: {DEFAULT_SUMMARY})",
    )
    parser.add_argument(
        "--max-sample-rows",
        type=int,
        default=200,
        help="Maximum number of per-sample rows kept per tool in output JSON.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute report and print summary without writing files.",
    )
    args = parser.parse_args()

    input_path = args.input.resolve()
    output_path = args.output.resolve()
    summary_path = args.summary.resolve()
    if not input_path.exists():
        print(f"Input report not found: {input_path}")
        return 2

    raw = _load_json(input_path)
    tool_names = _collect_tool_names(raw)
    tools = {tool_name: _compare_tool(raw, tool_name) for tool_name in tool_names}
    for tool_payload in tools.values():
        samples = tool_payload.get("samples", [])
        if isinstance(samples, list) and args.max_sample_rows > 0:
            tool_payload["samples"] = samples[: args.max_sample_rows]

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "contract_version": "odf-reference-v1",
        "input_report": str(args.input),
        "tool_order": tool_names,
        "tools": tools,
    }

    markdown = _build_markdown(report)
    print(markdown)

    if not args.dry_run:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
        summary_path.write_text(markdown, encoding="utf-8")
        print(f"JSON report written to {output_path}")
        print(f"Summary report written to {summary_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
