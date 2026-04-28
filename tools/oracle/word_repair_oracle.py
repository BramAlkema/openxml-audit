"""Spec 010 oracle driver: roundtrip generated property-ordering scenarios
through Word and emit a JSON oracle baseline.

Usage:

    # Run the full CT_TrPr matrix (baseline + pairwise swaps + full reverse).
    # Writes tools/oracle/baselines/word_trpr_pairwise.json.
    python -m tools.oracle.word_repair_oracle trpr

This is developer-machine infrastructure. See `tools/oracle/README.md`
for setup. Spec: `specs/011-word-roundtrip-oracle.md` Phase 2.
"""

from __future__ import annotations

import argparse
import json
import platform
import shutil
import sys
import time
import traceback
import uuid
from dataclasses import asdict, dataclass
from datetime import UTC, datetime
from pathlib import Path

from tools.oracle.diff import diff_property_fragment, extract_first
from tools.oracle.scenarios.property_ordering import (
    TRPR_CANONICAL_CORE,
    ScenarioSpec,
    materialize_trpr_scenario,
    trpr_baseline,
    trpr_full_reverse,
    trpr_pairwise_swaps,
)
from tools.oracle.word_roundtrip import RoundtripError, roundtrip
from tools.oracle.word_window import word_version

from openxml_audit.namespaces import WORDPROCESSINGML

REPO_ROOT = Path(__file__).resolve().parents[2]
BASELINES_DIR = Path(__file__).resolve().parent / "baselines"


@dataclass
class ScenarioResult:
    id: str
    parent_local: str
    description: str
    input_children: list[str]
    output_children: list[str] | None
    repair_dialog_seen: bool
    repair_dialog_text: str | None
    verdict: str  # "preserved" | "repaired" | "missing" | "engine_error"
    elapsed_seconds: float
    error: str | None


def _close_word_quietly() -> None:
    import contextlib
    import subprocess

    with contextlib.suppress(Exception):
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Word" to quit saving no'],
            timeout=10.0,
            check=False,
            capture_output=True,
        )


def _run_scenario(
    spec: ScenarioSpec, scenario_dir: Path, materialize: callable
) -> ScenarioResult:
    """Materialize one scenario, roundtrip it, classify the outcome."""
    started = time.monotonic()
    try:
        docx_path = scenario_dir / f"{spec.id}.docx"
        materialize(spec, docx_path)
        result = roundtrip(docx_path, timeout=60.0)
        input_children = extract_first(str(result.input_path), spec.parent_local, WORDPROCESSINGML)
        output_children = extract_first(str(result.output_path), spec.parent_local, WORDPROCESSINGML)
        diff = diff_property_fragment(spec.parent_local, input_children, output_children)
        return ScenarioResult(
            id=spec.id,
            parent_local=spec.parent_local,
            description=spec.description,
            input_children=list(spec.input_children),
            output_children=output_children,
            repair_dialog_seen=result.repair_dialog_seen,
            repair_dialog_text=result.repair_dialog_text,
            verdict=diff.verdict,
            elapsed_seconds=result.elapsed_seconds,
            error=None,
        )
    except (RoundtripError, Exception) as exc:  # noqa: BLE001 — record any failure as a result
        return ScenarioResult(
            id=spec.id,
            parent_local=spec.parent_local,
            description=spec.description,
            input_children=list(spec.input_children),
            output_children=None,
            repair_dialog_seen=False,
            repair_dialog_text=None,
            verdict="engine_error",
            elapsed_seconds=time.monotonic() - started,
            error=f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}",
        )


def _summarize(results: list[ScenarioResult]) -> dict:
    summary = {
        "total": len(results),
        "preserved": sum(1 for r in results if r.verdict == "preserved"),
        "repaired": sum(1 for r in results if r.verdict == "repaired"),
        "missing": sum(1 for r in results if r.verdict == "missing"),
        "engine_error": sum(1 for r in results if r.verdict == "engine_error"),
        "any_repair_dialog": any(r.repair_dialog_seen for r in results),
        "repair_dialogs_count": sum(1 for r in results if r.repair_dialog_seen),
    }
    return summary


def run_trpr_oracle(*, output_path: Path | None = None) -> dict:
    """Run the full CT_TrPr scenario matrix and emit the JSON report."""
    canonical = TRPR_CANONICAL_CORE
    scenarios: list[ScenarioSpec] = (
        [trpr_baseline(canonical)]
        + trpr_pairwise_swaps(canonical)
        + [trpr_full_reverse(canonical)]
    )

    work_root = Path.home() / "Documents" / ".word_oracle_runs" / f"trpr-batch-{uuid.uuid4().hex[:8]}"
    work_root.mkdir(parents=True, exist_ok=False)

    print(f"Running {len(scenarios)} CT_TrPr scenarios → {work_root}", file=sys.stderr)

    results: list[ScenarioResult] = []
    for i, spec in enumerate(scenarios, 1):
        print(f"  [{i}/{len(scenarios)}] {spec.id} ...", file=sys.stderr, end="", flush=True)
        result = _run_scenario(spec, work_root, materialize_trpr_scenario)
        results.append(result)
        marker = {"preserved": "·", "repaired": "!", "missing": "?", "engine_error": "X"}[result.verdict]
        dialog = " (dialog)" if result.repair_dialog_seen else ""
        print(f" {marker}{dialog} {result.elapsed_seconds:.1f}s", file=sys.stderr)

    report = {
        "constraint": "CT_TrPr",
        "spec_section": "ECMA-376 §17.4.79",
        "canonical_children_tested": list(canonical),
        "engine": "spec 011 Phase 1 (AppleScript open + close-with-save)",
        "word_version": word_version(),
        "platform": f"{platform.system()} {platform.release()}",
        "run_at": datetime.now(UTC).isoformat(timespec="seconds"),
        "summary": _summarize(results),
        "scenarios": [asdict(r) for r in results],
    }

    target = output_path or (BASELINES_DIR / "word_trpr_pairwise.json")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(report, indent=2) + "\n")
    print(f"\nWrote {target}", file=sys.stderr)
    print(f"Summary: {report['summary']}", file=sys.stderr)

    # Cleanup batch staging — committed baseline JSON has all the data we need
    shutil.rmtree(work_root, ignore_errors=True)
    _close_word_quietly()
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "constraint",
        choices=("trpr",),
        help="Which constraint type to run the oracle for. Phase 2 ships trpr only.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Output path for the JSON baseline (default: tools/oracle/baselines/word_<constraint>_pairwise.json).",
    )
    args = parser.parse_args()

    if args.constraint == "trpr":
        run_trpr_oracle(output_path=args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
