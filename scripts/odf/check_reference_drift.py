#!/usr/bin/env python3
"""Compare ODF reference comparison reports and enforce drift policy."""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

DEFAULT_BASELINE = Path("data/odf/reference_baseline/2026-03-09/mismatch_report.json")
DEFAULT_CURRENT = Path("reports/odf/reference_compare.json")
DEFAULT_OUTPUT = Path("reports/odf/reference_drift.json")
DEFAULT_SUMMARY = Path("reports/odf/reference_drift.md")
DEFAULT_POLICY = Path("data/odf/reference_baseline/2026-03-09/drift_policy.json")
DEFAULT_WAIVERS = Path("data/odf/reference_baseline/2026-03-09/waivers.json")

DEFAULT_POLICY_VALUES = {
    "max_only_python_growth": 0,
    "max_only_reference_growth": 0,
    "max_new_only_python_families": 0,
    "max_new_only_reference_families": 0,
    "max_compared_sample_drop": 0,
    "max_unavailable_samples": 0,
    "max_timeout_samples": 0,
    "max_error_samples": 0,
    "required_tools": ["odf_toolkit", "opf"],
}

WAIVER_KINDS = {
    "only_python_growth",
    "only_reference_growth",
    "new_only_python_family",
    "new_only_reference_family",
    "samples_compared_drop",
    "reference_unavailable",
    "reference_timeout",
    "reference_error",
}


@dataclass(frozen=True)
class Waiver:
    """Active waiver entry."""

    kind: str
    tool: str | None
    target: str | None
    owner: str
    reason: str
    expires: str


@dataclass(frozen=True)
class ToolSnapshot:
    """Tool-level normalized snapshot extracted from compare report."""

    name: str
    samples_total: int
    samples_compared: int
    only_python_total: int
    only_reference_total: int
    status_counts: dict[str, int]
    only_python_families: dict[str, int]
    only_reference_families: dict[str, int]


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _coerce_int(value: Any) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        try:
            return int(value)
        except ValueError:
            return 0
    return 0


def _coerce_policy_int(payload: dict[str, Any], key: str, default: int) -> int:
    value = payload.get(key, default)
    return _coerce_int(value)


def _normalize_family_counts(payload: Any) -> dict[str, int]:
    rows = payload if isinstance(payload, list) else []
    output: dict[str, int] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        key = row.get("family_group_key")
        if not isinstance(key, str) or not key:
            key = row.get("comparison_key")
        if not isinstance(key, str) or not key:
            key = row.get("description")
        if not isinstance(key, str) or not key:
            continue
        count = _coerce_int(row.get("count"))
        if count <= 0:
            continue
        output[key] = count
    return output


def _extract_tool_snapshot(report: dict[str, Any], tool_name: str) -> ToolSnapshot:
    tools = report.get("tools")
    if not isinstance(tools, dict):
        tools = {}
    tool_payload = tools.get(tool_name)
    if not isinstance(tool_payload, dict):
        tool_payload = {}

    status_counts_payload = tool_payload.get("reference_status_counts")
    if not isinstance(status_counts_payload, dict):
        status_counts_payload = {}
    status_counts: dict[str, int] = {
        str(key): _coerce_int(value) for key, value in status_counts_payload.items()
    }

    issue_totals = tool_payload.get("issue_totals")
    if not isinstance(issue_totals, dict):
        issue_totals = {}

    mismatch = tool_payload.get("mismatch_families")
    if not isinstance(mismatch, dict):
        mismatch = {}

    return ToolSnapshot(
        name=tool_name,
        samples_total=_coerce_int(tool_payload.get("samples_total")),
        samples_compared=_coerce_int(tool_payload.get("samples_compared")),
        only_python_total=_coerce_int(issue_totals.get("only_python")),
        only_reference_total=_coerce_int(issue_totals.get("only_reference")),
        status_counts=status_counts,
        only_python_families=_normalize_family_counts(mismatch.get("only_python")),
        only_reference_families=_normalize_family_counts(mismatch.get("only_reference")),
    )


def _collect_tool_names(*reports: dict[str, Any]) -> list[str]:
    names: set[str] = set()
    for report in reports:
        tool_order = report.get("tool_order")
        if isinstance(tool_order, list):
            for name in tool_order:
                if isinstance(name, str) and name:
                    names.add(name)
        tools = report.get("tools")
        if isinstance(tools, dict):
            for name in tools:
                names.add(str(name))
    return sorted(names)


def _load_policy(path: Path | None) -> tuple[dict[str, Any], list[str]]:
    policy: dict[str, Any] = dict(DEFAULT_POLICY_VALUES)
    warnings: list[str] = []

    if path is None:
        return policy, warnings
    if not path.exists():
        warnings.append(f"Policy file not found, defaults applied: {path}")
        return policy, warnings

    payload = _load_json(path)
    for key in (
        "max_only_python_growth",
        "max_only_reference_growth",
        "max_new_only_python_families",
        "max_new_only_reference_families",
        "max_compared_sample_drop",
        "max_unavailable_samples",
        "max_timeout_samples",
        "max_error_samples",
    ):
        policy[key] = _coerce_policy_int(payload, key, int(policy[key]))

    required_tools = payload.get("required_tools", policy["required_tools"])
    if isinstance(required_tools, list):
        normalized_tools = [tool for tool in required_tools if isinstance(tool, str) and tool]
        if normalized_tools:
            policy["required_tools"] = normalized_tools
        else:
            warnings.append("Policy required_tools is empty; defaults applied.")
    else:
        warnings.append("Policy required_tools is invalid; defaults applied.")

    return policy, warnings


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
        tool = row.get("tool")
        target = row.get("target")
        owner = row.get("owner")
        reason = row.get("reason")
        expires = row.get("expires")

        if not isinstance(kind, str) or kind not in WAIVER_KINDS:
            warnings.append(f"Waiver entry {idx} has invalid kind: {kind!r}.")
            continue
        if tool is not None and (not isinstance(tool, str) or not tool.strip()):
            warnings.append(f"Waiver entry {idx} has invalid tool.")
            continue
        if target is not None and (not isinstance(target, str) or not target.strip()):
            warnings.append(f"Waiver entry {idx} has invalid target.")
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
        if kind in {"new_only_python_family", "new_only_reference_family"} and (
            target is None or not isinstance(target, str) or not target.strip()
        ):
            warnings.append(f"Waiver entry {idx} kind={kind} requires target.")
            continue

        try:
            expires_date = date.fromisoformat(expires)
        except ValueError:
            warnings.append(f"Waiver entry {idx} has invalid expires date: {expires!r}.")
            continue
        if expires_date < today:
            warnings.append(
                f"Waiver expired and ignored: kind={kind} tool={tool!r} "
                f"target={target!r} expires={expires}."
            )
            continue

        active.append(
            Waiver(
                kind=kind,
                tool=tool.strip() if isinstance(tool, str) and tool.strip() else None,
                target=target.strip() if isinstance(target, str) and target.strip() else None,
                owner=owner.strip(),
                reason=reason.strip(),
                expires=expires,
            )
        )
    return active, warnings


def _find_waiver(
    waivers: list[Waiver],
    *,
    kind: str,
    tool: str,
    target: str | None = None,
) -> Waiver | None:
    for waiver in waivers:
        if waiver.kind != kind:
            continue
        if waiver.tool is not None and waiver.tool != tool:
            continue
        if waiver.target is not None and waiver.target != target:
            continue
        return waiver
    return None


def _serialize_waiver(waiver: Waiver) -> dict[str, str]:
    return {
        "owner": waiver.owner,
        "reason": waiver.reason,
        "expires": waiver.expires,
    }


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
        row = {
            "family_group_key": key,
            "baseline": before,
            "current": after,
            "delta": delta,
        }
        if before == 0 and after > 0:
            new_rows.append(row)
        if delta > 0:
            regressed_rows.append(row)
        elif delta < 0:
            improved_rows.append(row)

    new_rows.sort(key=lambda row: int(row["current"]), reverse=True)
    regressed_rows.sort(key=lambda row: int(row["delta"]), reverse=True)
    improved_rows.sort(key=lambda row: int(row["delta"]))
    return new_rows, regressed_rows, improved_rows


def _build_markdown(
    comparison: dict[str, Any], failures: list[str], waived_failures: list[str]
) -> str:
    lines = ["# ODF Reference Drift Summary", ""]
    lines.append(f"- Baseline: {comparison['baseline_path']}")
    lines.append(f"- Current: {comparison['current_path']}")
    lines.append(f"- Tools: {', '.join(comparison.get('tool_order', []))}")
    lines.append("")

    if failures:
        lines.append("## Gate result: FAILED")
        for failure in failures:
            lines.append(f"- {failure}")
    else:
        lines.append("## Gate result: PASSED")

    if waived_failures:
        lines.append("")
        lines.append("## Waived conditions")
        for detail in waived_failures:
            lines.append(f"- {detail}")

    for tool_name in comparison.get("tool_order", []):
        tool = comparison.get("tools", {}).get(tool_name)
        if not isinstance(tool, dict):
            continue
        lines.append("")
        lines.append(f"## {tool_name}")
        lines.append(
            f"- Samples compared: {tool['baseline']['samples_compared']} -> "
            f"{tool['current']['samples_compared']} "
            f"(delta {tool['deltas']['samples_compared']:+d})"
        )
        lines.append(
            f"- only_python mismatches: {tool['baseline']['only_python_total']} -> "
            f"{tool['current']['only_python_total']} "
            f"(delta {tool['deltas']['only_python_total']:+d})"
        )
        lines.append(
            f"- only_reference mismatches: {tool['baseline']['only_reference_total']} -> "
            f"{tool['current']['only_reference_total']} "
            f"(delta {tool['deltas']['only_reference_total']:+d})"
        )
        status_delta = tool.get("deltas", {}).get("status_counts", {})
        if isinstance(status_delta, dict):
            lines.append(f"- Status deltas: {status_delta}")

        family_summary = tool.get("new_family_counts", {})
        if isinstance(family_summary, dict):
            lines.append(
                "- New families (unwaived/waived): "
                f"only_python={family_summary.get('only_python_unwaived', 0)}/"
                f"{family_summary.get('only_python_waived', 0)}, "
                f"only_reference={family_summary.get('only_reference_unwaived', 0)}/"
                f"{family_summary.get('only_reference_waived', 0)}"
            )

        regressed = tool.get("families", {}).get("only_python", {}).get("regressed", [])
        if isinstance(regressed, list) and regressed:
            lines.append("")
            lines.append("### Top only-python regressions")
            for row in regressed[:5]:
                if not isinstance(row, dict):
                    continue
                lines.append(
                    f"- {row['delta']}x {row['family_group_key']} "
                    f"(baseline {row['baseline']}, current {row['current']})"
                )

        regressed_reference = tool.get("families", {}).get("only_reference", {}).get(
            "regressed", []
        )
        if isinstance(regressed_reference, list) and regressed_reference:
            lines.append("")
            lines.append("### Top only-reference regressions")
            for row in regressed_reference[:5]:
                if not isinstance(row, dict):
                    continue
                lines.append(
                    f"- {row['delta']}x {row['family_group_key']} "
                    f"(baseline {row['baseline']}, current {row['current']})"
                )

    warnings = comparison.get("waiver_warnings", [])
    if isinstance(warnings, list) and warnings:
        lines.append("")
        lines.append("## Waiver warnings")
        for warning in warnings:
            lines.append(f"- {warning}")

    policy_warnings = comparison.get("policy_warnings", [])
    if isinstance(policy_warnings, list) and policy_warnings:
        lines.append("")
        lines.append("## Policy warnings")
        for warning in policy_warnings:
            lines.append(f"- {warning}")

    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Check ODF reference drift against baseline.")
    parser.add_argument(
        "--baseline",
        type=Path,
        default=DEFAULT_BASELINE,
        help=f"Baseline compare report path (default: {DEFAULT_BASELINE})",
    )
    parser.add_argument(
        "--current",
        type=Path,
        default=DEFAULT_CURRENT,
        help=f"Current compare report path (default: {DEFAULT_CURRENT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output JSON path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=DEFAULT_SUMMARY,
        help=f"Output markdown summary path (default: {DEFAULT_SUMMARY})",
    )
    parser.add_argument(
        "--policy",
        type=Path,
        default=DEFAULT_POLICY,
        help=f"Threshold policy JSON path (default: {DEFAULT_POLICY})",
    )
    parser.add_argument(
        "--waivers",
        type=Path,
        default=DEFAULT_WAIVERS,
        help=f"Optional waiver file path (default: {DEFAULT_WAIVERS})",
    )
    parser.add_argument("--max-only-python-growth", type=int, default=None)
    parser.add_argument("--max-only-reference-growth", type=int, default=None)
    parser.add_argument("--max-new-only-python-families", type=int, default=None)
    parser.add_argument("--max-new-only-reference-families", type=int, default=None)
    parser.add_argument("--max-compared-sample-drop", type=int, default=None)
    parser.add_argument("--max-unavailable-samples", type=int, default=None)
    parser.add_argument("--max-timeout-samples", type=int, default=None)
    parser.add_argument("--max-error-samples", type=int, default=None)
    parser.add_argument("--dry-run", action="store_true", help="Do not write output files.")
    args = parser.parse_args()

    baseline_path = args.baseline.resolve()
    current_path = args.current.resolve()
    if not baseline_path.exists():
        print(f"Baseline compare report not found: {baseline_path}")
        return 2
    if not current_path.exists():
        print(f"Current compare report not found: {current_path}")
        return 2

    policy, policy_warnings = _load_policy(args.policy.resolve() if args.policy else None)
    if args.max_only_python_growth is not None:
        policy["max_only_python_growth"] = args.max_only_python_growth
    if args.max_only_reference_growth is not None:
        policy["max_only_reference_growth"] = args.max_only_reference_growth
    if args.max_new_only_python_families is not None:
        policy["max_new_only_python_families"] = args.max_new_only_python_families
    if args.max_new_only_reference_families is not None:
        policy["max_new_only_reference_families"] = args.max_new_only_reference_families
    if args.max_compared_sample_drop is not None:
        policy["max_compared_sample_drop"] = args.max_compared_sample_drop
    if args.max_unavailable_samples is not None:
        policy["max_unavailable_samples"] = args.max_unavailable_samples
    if args.max_timeout_samples is not None:
        policy["max_timeout_samples"] = args.max_timeout_samples
    if args.max_error_samples is not None:
        policy["max_error_samples"] = args.max_error_samples

    waivers, waiver_warnings = _load_waivers(args.waivers.resolve() if args.waivers else None)
    baseline_report = _load_json(baseline_path)
    current_report = _load_json(current_path)

    tool_names = _collect_tool_names(baseline_report, current_report)
    required_tools = [
        tool
        for tool in policy.get("required_tools", [])
        if isinstance(tool, str) and tool.strip()
    ]
    if required_tools:
        tool_names = sorted(set(tool_names) | set(required_tools))

    comparison_tools: dict[str, Any] = {}
    failures: list[str] = []
    waived_failures: list[str] = []

    for tool_name in tool_names:
        baseline = _extract_tool_snapshot(baseline_report, tool_name)
        current = _extract_tool_snapshot(current_report, tool_name)

        only_python_new, only_python_regressed, only_python_improved = _family_deltas(
            baseline.only_python_families,
            current.only_python_families,
        )
        only_reference_new, only_reference_regressed, only_reference_improved = _family_deltas(
            baseline.only_reference_families,
            current.only_reference_families,
        )

        only_python_new_waived: list[dict[str, Any]] = []
        only_python_new_unwaived: list[dict[str, Any]] = []
        for row in only_python_new:
            waiver = _find_waiver(
                waivers,
                kind="new_only_python_family",
                tool=tool_name,
                target=str(row["family_group_key"]),
            )
            if waiver is None:
                only_python_new_unwaived.append(row)
            else:
                only_python_new_waived.append({**row, "waiver": _serialize_waiver(waiver)})

        only_reference_new_waived: list[dict[str, Any]] = []
        only_reference_new_unwaived: list[dict[str, Any]] = []
        for row in only_reference_new:
            waiver = _find_waiver(
                waivers,
                kind="new_only_reference_family",
                tool=tool_name,
                target=str(row["family_group_key"]),
            )
            if waiver is None:
                only_reference_new_unwaived.append(row)
            else:
                only_reference_new_waived.append({**row, "waiver": _serialize_waiver(waiver)})

        samples_compared_delta = current.samples_compared - baseline.samples_compared
        only_python_growth = current.only_python_total - baseline.only_python_total
        only_reference_growth = current.only_reference_total - baseline.only_reference_total
        status_deltas = {
            key: current.status_counts.get(key, 0) - baseline.status_counts.get(key, 0)
            for key in sorted(set(baseline.status_counts) | set(current.status_counts))
        }
        compared_drop = baseline.samples_compared - current.samples_compared
        unavailable_current = current.status_counts.get("unavailable", 0)
        timeout_current = current.status_counts.get("timeout", 0)
        error_current = current.status_counts.get("error", 0)

        comparison_tools[tool_name] = {
            "baseline": {
                "samples_total": baseline.samples_total,
                "samples_compared": baseline.samples_compared,
                "only_python_total": baseline.only_python_total,
                "only_reference_total": baseline.only_reference_total,
                "status_counts": baseline.status_counts,
            },
            "current": {
                "samples_total": current.samples_total,
                "samples_compared": current.samples_compared,
                "only_python_total": current.only_python_total,
                "only_reference_total": current.only_reference_total,
                "status_counts": current.status_counts,
            },
            "deltas": {
                "samples_compared": samples_compared_delta,
                "only_python_total": only_python_growth,
                "only_reference_total": only_reference_growth,
                "status_counts": status_deltas,
            },
            "families": {
                "only_python": {
                    "new": only_python_new,
                    "new_waived": only_python_new_waived,
                    "new_unwaived": only_python_new_unwaived,
                    "regressed": only_python_regressed,
                    "improved": only_python_improved,
                },
                "only_reference": {
                    "new": only_reference_new,
                    "new_waived": only_reference_new_waived,
                    "new_unwaived": only_reference_new_unwaived,
                    "regressed": only_reference_regressed,
                    "improved": only_reference_improved,
                },
            },
            "new_family_counts": {
                "only_python_total": len(only_python_new),
                "only_python_waived": len(only_python_new_waived),
                "only_python_unwaived": len(only_python_new_unwaived),
                "only_reference_total": len(only_reference_new),
                "only_reference_waived": len(only_reference_new_waived),
                "only_reference_unwaived": len(only_reference_new_unwaived),
            },
        }

        max_only_python_growth = int(policy["max_only_python_growth"])
        if only_python_growth > max_only_python_growth:
            waiver = _find_waiver(waivers, kind="only_python_growth", tool=tool_name)
            if waiver is not None:
                waived_failures.append(
                    f"{tool_name}: only_python_growth waiver applied "
                    f"(owner={waiver.owner}, expires={waiver.expires})."
                )
            else:
                failures.append(
                    f"{tool_name}: only_python growth exceeded threshold "
                    f"({only_python_growth} > {max_only_python_growth})."
                )

        max_only_reference_growth = int(policy["max_only_reference_growth"])
        if only_reference_growth > max_only_reference_growth:
            waiver = _find_waiver(waivers, kind="only_reference_growth", tool=tool_name)
            if waiver is not None:
                waived_failures.append(
                    f"{tool_name}: only_reference_growth waiver applied "
                    f"(owner={waiver.owner}, expires={waiver.expires})."
                )
            else:
                failures.append(
                    f"{tool_name}: only_reference growth exceeded threshold "
                    f"({only_reference_growth} > {max_only_reference_growth})."
                )

        max_new_only_python = int(policy["max_new_only_python_families"])
        if len(only_python_new_unwaived) > max_new_only_python:
            failures.append(
                f"{tool_name}: new only-python family count exceeded threshold "
                f"({len(only_python_new_unwaived)} > {max_new_only_python})."
            )

        max_new_only_reference = int(policy["max_new_only_reference_families"])
        if len(only_reference_new_unwaived) > max_new_only_reference:
            failures.append(
                f"{tool_name}: new only-reference family count exceeded threshold "
                f"({len(only_reference_new_unwaived)} > {max_new_only_reference})."
            )

        max_compared_drop = int(policy["max_compared_sample_drop"])
        if compared_drop > max_compared_drop:
            waiver = _find_waiver(waivers, kind="samples_compared_drop", tool=tool_name)
            if waiver is not None:
                waived_failures.append(
                    f"{tool_name}: samples_compared_drop waiver applied "
                    f"(owner={waiver.owner}, expires={waiver.expires})."
                )
            else:
                failures.append(
                    f"{tool_name}: samples compared drop exceeded threshold "
                    f"({compared_drop} > {max_compared_drop})."
                )

        max_unavailable = int(policy["max_unavailable_samples"])
        if unavailable_current > max_unavailable:
            waiver = _find_waiver(waivers, kind="reference_unavailable", tool=tool_name)
            if waiver is not None:
                waived_failures.append(
                    f"{tool_name}: reference_unavailable waiver applied "
                    f"(owner={waiver.owner}, expires={waiver.expires})."
                )
            else:
                failures.append(
                    f"{tool_name}: unavailable sample count exceeded threshold "
                    f"({unavailable_current} > {max_unavailable}); "
                    "check reference-validator command wiring."
                )

        max_timeout = int(policy["max_timeout_samples"])
        if timeout_current > max_timeout:
            waiver = _find_waiver(waivers, kind="reference_timeout", tool=tool_name)
            if waiver is not None:
                waived_failures.append(
                    f"{tool_name}: reference_timeout waiver applied "
                    f"(owner={waiver.owner}, expires={waiver.expires})."
                )
            else:
                failures.append(
                    f"{tool_name}: timeout sample count exceeded threshold "
                    f"({timeout_current} > {max_timeout})."
                )

        max_error = int(policy["max_error_samples"])
        if error_current > max_error:
            waiver = _find_waiver(waivers, kind="reference_error", tool=tool_name)
            if waiver is not None:
                waived_failures.append(
                    f"{tool_name}: reference_error waiver applied "
                    f"(owner={waiver.owner}, expires={waiver.expires})."
                )
            else:
                failures.append(
                    f"{tool_name}: error sample count exceeded threshold "
                    f"({error_current} > {max_error})."
                )

    current_tools_payload = current_report.get("tools")
    if not isinstance(current_tools_payload, dict):
        current_tools_payload = {}
    if required_tools:
        for required_tool in required_tools:
            if required_tool not in current_tools_payload:
                failures.append(
                    f"Required tool missing in current compare payload: {required_tool}."
                )

    comparison = {
        "baseline_path": str(baseline_path),
        "current_path": str(current_path),
        "tool_order": tool_names,
        "policy": policy,
        "tools": comparison_tools,
        "policy_warnings": policy_warnings,
        "waiver_warnings": waiver_warnings,
    }

    comparison["gate"] = {
        "passed": not failures,
        "failures": failures,
        "waived_conditions": waived_failures,
    }

    print(f"Checked tools: {', '.join(tool_names)}")
    print(f"Gate passed: {comparison['gate']['passed']}")
    if failures:
        print("Failures:")
        for failure in failures:
            print(f"- {failure}")
    if waived_failures:
        print("Waived conditions:")
        for detail in waived_failures:
            print(f"- {detail}")

    summary_content = _build_markdown(
        comparison=comparison,
        failures=failures,
        waived_failures=waived_failures,
    )
    print("")
    print(summary_content)

    if not args.dry_run:
        output_path = args.output.resolve()
        summary_path = args.summary.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(comparison, indent=2), encoding="utf-8")
        summary_path.write_text(summary_content, encoding="utf-8")
        print(f"Wrote comparison: {output_path}")
        print(f"Wrote summary: {summary_path}")

    return 0 if not failures else 1


if __name__ == "__main__":
    raise SystemExit(main())
