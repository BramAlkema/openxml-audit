"""Tests for the self-parity prototype (Spec 027 / 0.7.1).

The snapshot generator + comparator under `scripts/parity/` are the
prototype shipped as advisory CI in 0.7.1. 0.8.0 promotes the
comparator to the blocking sovereign gate.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from parity.compare_self_parity import (  # noqa: E402
    compare,
    evaluate_gate,
    render_summary_md,
)


def _make_inventory(entries: list[tuple[str, int, str]]) -> dict:
    """Build a synthetic family_inventory dict matching the snapshot shape."""
    return {
        key: {
            "family_key": key,
            "count": count,
            "source_class": cls,
            "first_part": "",
            "first_path_sample": "",
            "first_description_sample": "",
        }
        for key, count, cls in entries
    }


def _make_snapshot(entries: list[tuple[str, int, str]]) -> dict:
    inventory = _make_inventory(entries)
    return {
        "schema_version": 1,
        "validator_version": "test",
        "files_root": "test",
        "manifest": "test",
        "file_count": 1,
        "validation_runs": 1,
        "total_findings": sum(e[1] for e in entries),
        "by_source_class": {},
        "family_inventory": inventory,
    }


def test_compare_reports_zero_drift_when_snapshots_identical() -> None:
    snap = _make_snapshot(
        [
            ("Sch_X|schema|/word/document.xml|/document[1]|attr", 5, "sdk_proxy"),
            ("Sem_Y|semantic|/word/document.xml|/document[1]|word_compat", 2, "word_app_compat"),
        ]
    )
    report = compare(snap, snap)
    assert report.new_families == []
    assert report.missing_families == []
    assert report.count_drift == []
    assert report.baseline_total == 7
    assert report.current_total == 7


def test_compare_detects_new_families() -> None:
    base = _make_snapshot([("A", 1, "sdk_proxy")])
    cur = _make_snapshot([("A", 1, "sdk_proxy"), ("B", 3, "sdk_proxy")])
    report = compare(base, cur)
    assert len(report.new_families) == 1
    diff = report.new_families[0]
    assert diff.family_key == "B"
    assert diff.baseline_count == 0
    assert diff.current_count == 3
    assert diff.delta == 3
    assert diff.source_class == "sdk_proxy"


def test_compare_detects_missing_families() -> None:
    base = _make_snapshot([("A", 1, "sdk_proxy"), ("B", 3, "word_app_compat")])
    cur = _make_snapshot([("A", 1, "sdk_proxy")])
    report = compare(base, cur)
    assert len(report.missing_families) == 1
    diff = report.missing_families[0]
    assert diff.family_key == "B"
    assert diff.baseline_count == 3
    assert diff.current_count == 0
    assert diff.delta == -3
    assert diff.source_class == "word_app_compat"


def test_compare_detects_count_drift() -> None:
    base = _make_snapshot([("A", 1, "sdk_proxy")])
    cur = _make_snapshot([("A", 5, "sdk_proxy")])
    report = compare(base, cur)
    assert len(report.count_drift) == 1
    diff = report.count_drift[0]
    assert diff.family_key == "A"
    assert diff.baseline_count == 1
    assert diff.current_count == 5
    assert diff.delta == 4


def test_evaluate_gate_passes_under_strict_no_drift() -> None:
    snap = _make_snapshot([("A", 1, "sdk_proxy")])
    report = compare(snap, snap)
    verdict = evaluate_gate(report)
    assert verdict["ok"] is True
    assert verdict["failures"] == []


def test_evaluate_gate_fails_on_new_family() -> None:
    base = _make_snapshot([])
    cur = _make_snapshot([("NEW", 1, "sdk_proxy")])
    report = compare(base, cur)
    verdict = evaluate_gate(report, max_new_families=0)
    assert verdict["ok"] is False
    assert any("new families" in f for f in verdict["failures"])


def test_evaluate_gate_passes_when_threshold_allows() -> None:
    base = _make_snapshot([])
    cur = _make_snapshot([("NEW", 1, "sdk_proxy")])
    report = compare(base, cur)
    verdict = evaluate_gate(report, max_new_families=5)
    assert verdict["ok"] is True


def test_evaluate_gate_uses_absolute_drift_total() -> None:
    """Count drift is summed as absolute deltas — a +3/-3 in two
    different families is 6 total drift, not 0. Catches "we lost a
    finding here, gained one there" net-zero scenarios that look
    clean by total count."""
    base = _make_snapshot([("A", 5, "sdk_proxy"), ("B", 5, "sdk_proxy")])
    cur = _make_snapshot([("A", 8, "sdk_proxy"), ("B", 2, "sdk_proxy")])
    report = compare(base, cur)
    verdict = evaluate_gate(report, max_count_drift_total=0)
    assert verdict["ok"] is False
    assert verdict["counts"]["count_drift_total"] == 6


def test_render_summary_md_includes_verdict_and_diffs() -> None:
    base = _make_snapshot([("A", 1, "sdk_proxy")])
    cur = _make_snapshot([("A", 1, "sdk_proxy"), ("B", 1, "word_app_compat")])
    report = compare(base, cur)
    report.gate = evaluate_gate(report)
    md = render_summary_md(report)
    assert "FAIL" in md  # one new family, default threshold zero
    assert "B" in md
    assert "word_app_compat" in md
    assert "Self-parity comparison" in md


def test_baseline_against_self_passes_via_real_baseline() -> None:
    """The committed v0.7.1 baseline must compare clean against itself.
    This locks the schema shape and prevents accidental field renames
    from breaking the comparator silently."""
    baseline_path = REPO_ROOT / "data/corpus/self_parity_baseline/v0.7.1/snapshot.json"
    if not baseline_path.exists():
        pytest.skip("v0.7.1 baseline not committed yet")
    baseline = json.loads(baseline_path.read_text())
    report = compare(baseline, baseline)
    assert report.new_families == []
    assert report.missing_families == []
    assert report.count_drift == []
