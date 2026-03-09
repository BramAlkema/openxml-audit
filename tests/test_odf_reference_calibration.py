from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]


def _run_command(args: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        args,
        cwd=REPO_ROOT,
        check=False,
        capture_output=True,
        text=True,
    )


def test_compare_reference_results_groups_cross_tool_families(tmp_path: Path) -> None:
    input_path = tmp_path / "reference_runs.json"
    output_path = tmp_path / "reference_compare.json"
    summary_path = tmp_path / "reference_compare.md"

    input_payload = {
        "samples": [
            {
                "id": "sample-1",
                "runs": {
                    "python": {
                        "status": "ok",
                        "issues": [
                            {
                                "comparison_key": "error|missing required style declaration",
                                "description": "missing required style declaration",
                                "category": "schema",
                            }
                        ],
                    },
                    "odf_toolkit": {"status": "ok", "issues": []},
                    "opf": {"status": "ok", "issues": []},
                },
            }
        ]
    }
    input_path.write_text(json.dumps(input_payload), encoding="utf-8")

    completed = _run_command(
        [
            sys.executable,
            "scripts/odf/compare_reference_results.py",
            "--input",
            str(input_path),
            "--output",
            str(output_path),
            "--summary",
            str(summary_path),
        ]
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr

    payload = json.loads(output_path.read_text(encoding="utf-8"))
    grouped = payload.get("cross_tool_families", {}).get("only_python", [])
    assert isinstance(grouped, list) and grouped
    top_row = grouped[0]
    assert top_row["count"] == 2
    assert top_row["tools"] == {"odf_toolkit": 1, "opf": 1}


def test_check_reference_drift_supports_tool_scoped_waiver(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.json"
    current_path = tmp_path / "current.json"
    policy_path = tmp_path / "policy.json"
    waivers_path = tmp_path / "waivers.json"
    output_path = tmp_path / "drift.json"
    summary_path = tmp_path / "drift.md"

    baseline_payload = {
        "tool_order": ["odf_toolkit", "opf"],
        "tools": {
            "odf_toolkit": {
                "samples_total": 1,
                "samples_compared": 1,
                "reference_status_counts": {"ok": 1},
                "issue_totals": {"only_python": 0, "only_reference": 0},
                "mismatch_families": {"only_python": [], "only_reference": []},
            },
            "opf": {
                "samples_total": 1,
                "samples_compared": 1,
                "reference_status_counts": {"ok": 1},
                "issue_totals": {"only_python": 0, "only_reference": 0},
                "mismatch_families": {"only_python": [], "only_reference": []},
            },
        },
    }
    current_payload = {
        "tool_order": ["odf_toolkit", "opf"],
        "tools": {
            "odf_toolkit": {
                "samples_total": 1,
                "samples_compared": 0,
                "reference_status_counts": {"unavailable": 1},
                "issue_totals": {"only_python": 0, "only_reference": 0},
                "mismatch_families": {"only_python": [], "only_reference": []},
            },
            "opf": {
                "samples_total": 1,
                "samples_compared": 1,
                "reference_status_counts": {"ok": 1},
                "issue_totals": {"only_python": 0, "only_reference": 0},
                "mismatch_families": {"only_python": [], "only_reference": []},
            },
        },
    }
    policy_payload = {
        "max_only_python_growth": 0,
        "max_only_reference_growth": 0,
        "max_new_only_python_families": 0,
        "max_new_only_reference_families": 0,
        "max_compared_sample_drop": 1,
        "max_unavailable_samples": 0,
        "max_timeout_samples": 0,
        "max_error_samples": 0,
        "required_tools": ["odf_toolkit", "opf"],
    }
    waivers_payload = {
        "waivers": [
            {
                "kind": "reference_unavailable",
                "tool": "odf_toolkit",
                "owner": "qa-maintainers",
                "reason": "Reference validator setup in progress",
                "expires": "2027-01-01",
            }
        ]
    }

    baseline_path.write_text(json.dumps(baseline_payload), encoding="utf-8")
    current_path.write_text(json.dumps(current_payload), encoding="utf-8")
    policy_path.write_text(json.dumps(policy_payload), encoding="utf-8")
    waivers_path.write_text(json.dumps(waivers_payload), encoding="utf-8")

    completed = _run_command(
        [
            sys.executable,
            "scripts/odf/check_reference_drift.py",
            "--baseline",
            str(baseline_path),
            "--current",
            str(current_path),
            "--policy",
            str(policy_path),
            "--waivers",
            str(waivers_path),
            "--output",
            str(output_path),
            "--summary",
            str(summary_path),
        ]
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr

    report = json.loads(output_path.read_text(encoding="utf-8"))
    assert report["gate"]["passed"] is True
    waived_conditions = report["gate"]["waived_conditions"]
    assert isinstance(waived_conditions, list)
    assert any("reference_unavailable waiver applied" in detail for detail in waived_conditions)
