#!/usr/bin/env python3
"""Build categorized ODF mismatch triage artifacts from comparison reports."""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_COMPARE = Path("reports/odf/reference_compare.json")
DEFAULT_RUNS = Path("reports/odf/reference_runs.json")
DEFAULT_OUTPUT = Path("reports/odf/mismatch_triage.md")


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _sample_category_counts(run_report: dict[str, Any]) -> dict[str, int]:
    counts: dict[str, int] = {}
    samples = run_report.get("samples")
    if not isinstance(samples, list):
        return counts
    for sample in samples:
        if not isinstance(sample, dict):
            continue
        category = sample.get("category")
        if not isinstance(category, str) or not category:
            continue
        counts[category] = counts.get(category, 0) + 1
    return dict(sorted(counts.items(), key=lambda row: row[0]))


def _sample_profile_counts(run_report: dict[str, Any]) -> dict[str, int]:
    counts: dict[str, int] = {}
    samples = run_report.get("samples")
    if not isinstance(samples, list):
        return counts
    for sample in samples:
        if not isinstance(sample, dict):
            continue
        profile = sample.get("profile")
        if not isinstance(profile, str) or not profile:
            continue
        counts[profile] = counts.get(profile, 0) + 1
    return dict(sorted(counts.items(), key=lambda row: row[0]))


def _build_markdown(compare_report: dict[str, Any], run_report: dict[str, Any] | None) -> str:
    lines = ["# ODF Mismatch Triage", ""]
    lines.append(f"- Generated at: {datetime.now(timezone.utc).isoformat()}")
    lines.append(f"- Compare input: {compare_report.get('input_report', '')}")
    lines.append(f"- Tools: {', '.join(compare_report.get('tool_order', []))}")

    if run_report is not None:
        lines.append(f"- Sample count: {run_report.get('sample_count', 0)}")
        lines.append(f"- Sample categories: {_sample_category_counts(run_report)}")
        lines.append(f"- Sample profiles: {_sample_profile_counts(run_report)}")
        lines.append(
            f"- Python issue categories: {run_report.get('python_issue_categories', {})}"
        )
    lines.append("")

    actions: list[str] = []
    if run_report is not None:
        runners = run_report.get("runners")
        if isinstance(runners, dict):
            for runner_name in ("odf_toolkit", "opf"):
                payload = runners.get(runner_name)
                if not isinstance(payload, dict):
                    continue
                status_counts = payload.get("status_counts")
                unavailable = 0
                if isinstance(status_counts, dict):
                    raw_unavailable = status_counts.get("unavailable", 0)
                    if isinstance(raw_unavailable, int):
                        unavailable = raw_unavailable
                if unavailable > 0:
                    template = payload.get("command_template")
                    if template is None:
                        if runner_name == "odf_toolkit":
                            actions.append(
                                "Configure ODF Toolkit command template "
                                "(`ODF_TOOLKIT_CMD`) in the calibration environment."
                            )
                        else:
                            actions.append(
                                "Configure OPF command template (`OPF_ODF_VALIDATOR_CMD`) "
                                "in the calibration environment."
                            )
                    else:
                        actions.append(
                            f"{runner_name} reported {unavailable} unavailable sample runs; "
                            "check validator startup and command template compatibility."
                        )

    cross_tool = compare_report.get("cross_tool_families")
    if isinstance(cross_tool, dict):
        only_python = cross_tool.get("only_python")
        if isinstance(only_python, list) and only_python:
            actions.append(
                "Prioritize top grouped only-python families for false-positive triage."
            )
        only_reference = cross_tool.get("only_reference")
        if isinstance(only_reference, list) and only_reference:
            actions.append(
                "Prioritize top grouped only-reference families for missing-rule coverage."
            )

    if actions:
        lines.append("## Actionable Summary")
        for action in dict.fromkeys(actions):
            lines.append(f"- {action}")
        lines.append("")

    tools = compare_report.get("tools", {})
    for tool_name in compare_report.get("tool_order", []):
        tool = tools.get(tool_name, {})
        if not isinstance(tool, dict):
            continue
        lines.append(f"## {tool_name}")
        lines.append(
            f"- Samples compared/skipped: {tool.get('samples_compared', 0)} / "
            f"{tool.get('samples_skipped', 0)}"
        )
        lines.append(
            f"- Reference status counts: {tool.get('reference_status_counts', {})}"
        )
        mismatch_categories = tool.get("mismatch_categories", {})
        lines.append(
            "- Mismatch categories: "
            f"only_python={mismatch_categories.get('only_python', {})}, "
            f"only_reference={mismatch_categories.get('only_reference', {})}"
        )

        only_python = tool.get("mismatch_families", {}).get("only_python", [])
        only_reference = tool.get("mismatch_families", {}).get("only_reference", [])
        if isinstance(only_python, list) and only_python:
            lines.append("")
            lines.append("### Top only-python families")
            for row in only_python[:10]:
                if isinstance(row, dict):
                    lines.append(f"- {row.get('count', 0)}x {row.get('description', '')}")
        if isinstance(only_reference, list) and only_reference:
            lines.append("")
            lines.append("### Top only-reference families")
            for row in only_reference[:10]:
                if isinstance(row, dict):
                    lines.append(f"- {row.get('count', 0)}x {row.get('description', '')}")

        compared = int(tool.get("samples_compared", 0) or 0)
        if compared == 0:
            lines.append("")
            lines.append("- No comparable runs yet for this tool (check command wiring/status).")
        lines.append("")

    if isinstance(cross_tool, dict):
        grouped_python = cross_tool.get("only_python")
        grouped_reference = cross_tool.get("only_reference")
        if isinstance(grouped_python, list) and grouped_python:
            lines.append("## Top grouped only-python families")
            for row in grouped_python[:10]:
                if isinstance(row, dict):
                    lines.append(
                        f"- {row.get('count', 0)}x {row.get('description', '')} "
                        f"(tools={row.get('tools', {})})"
                    )
            lines.append("")
        if isinstance(grouped_reference, list) and grouped_reference:
            lines.append("## Top grouped only-reference families")
            for row in grouped_reference[:10]:
                if isinstance(row, dict):
                    lines.append(
                        f"- {row.get('count', 0)}x {row.get('description', '')} "
                        f"(tools={row.get('tools', {})})"
                    )
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate categorized ODF mismatch triage markdown."
    )
    parser.add_argument(
        "--compare",
        type=Path,
        default=DEFAULT_COMPARE,
        help=f"Input compare report path (default: {DEFAULT_COMPARE})",
    )
    parser.add_argument(
        "--runs",
        type=Path,
        default=DEFAULT_RUNS,
        help=f"Optional run report path (default: {DEFAULT_RUNS})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output markdown path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print triage markdown without writing the output file.",
    )
    args = parser.parse_args()

    compare_path = args.compare.resolve()
    if not compare_path.exists():
        print(f"Compare report not found: {compare_path}")
        return 2
    compare_report = _load_json(compare_path)

    run_report: dict[str, Any] | None = None
    runs_path = args.runs.resolve()
    if runs_path.exists():
        run_report = _load_json(runs_path)

    markdown = _build_markdown(compare_report, run_report)
    print(markdown)

    if not args.dry_run:
        output_path = args.output.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(markdown, encoding="utf-8")
        print(f"Triage markdown written to {output_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
