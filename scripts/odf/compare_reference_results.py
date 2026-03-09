#!/usr/bin/env python3
"""Compare normalized Python ODF findings against reference validator runs."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_INPUT = Path("reports/odf/reference_runs.json")
DEFAULT_OUTPUT = Path("reports/odf/reference_compare.json")
DEFAULT_SUMMARY = Path("reports/odf/reference_compare.md")

_TOOL_NOISE_PATTERN = re.compile(
    r"\b(?:odf[ -]?toolkit|odf validator|opf|apache|validator|validation)\b",
    flags=re.IGNORECASE,
)
_FILE_NOISE_PATTERN = re.compile(
    r"\b[0-9A-Za-z_.-]+\.(?:xml|odt|ods|odp|fodt|fods|fodp)\b",
    flags=re.IGNORECASE,
)
_PATH_NOISE_PATTERN = re.compile(r"(?:[A-Za-z]+:)?/[0-9A-Za-z_.:/-]+")
_SEPARATOR_PATTERN = re.compile(r"[^0-9a-z<>]+", flags=re.IGNORECASE)


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


def _field_map(issues: Any, field: str) -> dict[str, str]:
    mapping: dict[str, str] = {}
    if not isinstance(issues, list):
        return mapping
    for issue in issues:
        if not isinstance(issue, dict):
            continue
        key = issue.get("comparison_key")
        if not isinstance(key, str) or not key:
            continue
        value = issue.get(field)
        if isinstance(value, str) and value and key not in mapping:
            mapping[key] = value
    return mapping


def _split_comparison_key(key: str) -> tuple[str, str]:
    severity, separator, description = key.partition("|")
    if not separator:
        return "error", key
    return severity.strip() or "error", description.strip() or key


def _normalize_cross_tool_description(value: str) -> str:
    normalized = value.lower()
    normalized = _TOOL_NOISE_PATTERN.sub(" ", normalized)
    normalized = _FILE_NOISE_PATTERN.sub("<file>", normalized)
    normalized = _PATH_NOISE_PATTERN.sub(" <path> ", normalized)
    normalized = _SEPARATOR_PATTERN.sub(" ", normalized)
    normalized = " ".join(normalized.split())
    return normalized or "unknown"


def _cross_tool_family_key(comparison_key: str) -> str:
    severity, description = _split_comparison_key(comparison_key)
    return f"{severity}|{_normalize_cross_tool_description(description)}"


def _sorted_family_rows(
    counter: Counter[str], descriptions: dict[str, str]
) -> list[dict[str, Any]]:
    rows = [
        {
            "comparison_key": key,
            "description": descriptions.get(key, key),
            "family_group_key": _cross_tool_family_key(key),
            "count": count,
        }
        for key, count in counter.items()
    ]
    rows.sort(key=lambda row: int(row["count"]), reverse=True)
    return rows


def _aggregate_cross_tool_families(tools: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    aggregated: dict[str, dict[str, dict[str, Any]]] = {
        "only_python": {},
        "only_reference": {},
    }
    for tool_name, tool_payload in tools.items():
        if not isinstance(tool_payload, dict):
            continue
        mismatch = tool_payload.get("mismatch_families")
        if not isinstance(mismatch, dict):
            continue
        for direction in ("only_python", "only_reference"):
            rows = mismatch.get(direction)
            if not isinstance(rows, list):
                continue
            for row in rows:
                if not isinstance(row, dict):
                    continue
                comparison_key = row.get("comparison_key")
                if not isinstance(comparison_key, str) or not comparison_key:
                    continue
                group_key = row.get("family_group_key")
                if not isinstance(group_key, str) or not group_key:
                    group_key = _cross_tool_family_key(comparison_key)
                count = row.get("count")
                if not isinstance(count, int):
                    continue

                group_bucket = aggregated[direction].setdefault(
                    group_key,
                    {
                        "family_group_key": group_key,
                        "description": row.get("description", comparison_key),
                        "count": 0,
                        "tools": {},
                    },
                )
                group_bucket["count"] = int(group_bucket["count"]) + count
                tool_counts = group_bucket["tools"]
                tool_counts[tool_name] = int(tool_counts.get(tool_name, 0)) + count

    output: dict[str, list[dict[str, Any]]] = {}
    for direction, buckets in aggregated.items():
        rows = list(buckets.values())
        rows.sort(key=lambda row: int(row.get("count", 0)), reverse=True)
        for row in rows:
            tools_map = row.get("tools")
            if isinstance(tools_map, dict):
                row["tools"] = {
                    key: int(tools_map[key])
                    for key in sorted(tools_map)
                    if isinstance(tools_map[key], int)
                }
        output[direction] = rows
    return output


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
    only_python_categories: Counter[str] = Counter()
    only_reference_categories: Counter[str] = Counter()
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
        python_descriptions = _field_map(python_run.get("issues"), "description")
        reference_descriptions = _field_map(reference_run.get("issues"), "description")
        python_categories = _field_map(python_run.get("issues"), "category")
        reference_categories = _field_map(reference_run.get("issues"), "category")

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
        for key, count in only_python.items():
            only_python_categories[python_categories.get(key, "unknown")] += count
        for key, count in only_reference.items():
            only_reference_categories[reference_categories.get(key, "unknown")] += count

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
        "mismatch_categories": {
            "only_python": dict(only_python_categories),
            "only_reference": dict(only_reference_categories),
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
        categories = tool.get("mismatch_categories", {})
        lines.append(
            "- Mismatch categories: "
            f"only_python={categories.get('only_python', {})}, "
            f"only_reference={categories.get('only_reference', {})}"
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

    cross_tool = report.get("cross_tool_families")
    if isinstance(cross_tool, dict):
        has_rows = False
        for direction in ("only_python", "only_reference"):
            rows = cross_tool.get(direction)
            if isinstance(rows, list) and rows:
                has_rows = True
                break
        if has_rows:
            lines.append("## Cross-tool grouped families")
        for direction, heading in (
            ("only_python", "Top grouped only-python families"),
            ("only_reference", "Top grouped only-reference families"),
        ):
            rows = cross_tool.get(direction)
            if not isinstance(rows, list) or not rows:
                continue
            lines.append("")
            lines.append(f"### {heading}")
            for row in rows[:10]:
                if not isinstance(row, dict):
                    continue
                lines.append(
                    f"- {row.get('count', 0)}x {row.get('description', '')} "
                    f"(tools={row.get('tools', {})})"
                )
        if has_rows:
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
        "--max-cross-tool-rows",
        type=int,
        default=200,
        help="Maximum grouped cross-tool families kept per mismatch direction.",
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
    cross_tool_families = _aggregate_cross_tool_families(tools)
    if args.max_cross_tool_rows > 0:
        for direction in ("only_python", "only_reference"):
            rows = cross_tool_families.get(direction)
            if isinstance(rows, list):
                cross_tool_families[direction] = rows[: args.max_cross_tool_rows]

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "contract_version": "odf-reference-v2",
        "input_report": str(args.input),
        "tool_order": tool_names,
        "tools": tools,
        "cross_tool_families": cross_tool_families,
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
