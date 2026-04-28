"""Spec 010 oracle driver: roundtrip generated property-ordering scenarios
through Word and emit a JSON oracle baseline.

Usage:

    # Run the full matrix for one constraint type.
    # Writes tools/oracle/baselines/word_<key>_pairwise.json.
    python -m tools.oracle.word_repair_oracle trpr
    python -m tools.oracle.word_repair_oracle tblpr
    python -m tools.oracle.word_repair_oracle tcpr
    python -m tools.oracle.word_repair_oracle sectpr

This is developer-machine infrastructure. See `tools/oracle/README.md`
for setup. Spec: `specs/011-word-roundtrip-oracle.md` Phase 2.
"""

from __future__ import annotations

import argparse
import contextlib
import json
import platform
import shutil
import subprocess
import sys
import time
import traceback
import uuid
from collections.abc import Callable
from dataclasses import asdict, dataclass
from datetime import UTC, datetime
from pathlib import Path

from tools.oracle.diff import diff_property_fragment, extract_first
from tools.oracle.scenarios.property_ordering import (
    MATRICES,
    OrderingMatrix,
    ScenarioSpec,
    all_scenarios,
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
    with contextlib.suppress(Exception):
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Word" to quit saving no'],
            timeout=10.0,
            check=False,
            capture_output=True,
        )


def _run_scenario(
    spec: ScenarioSpec,
    matrix: OrderingMatrix,
    scenario_dir: Path,
    materialize: Callable[[ScenarioSpec, OrderingMatrix, Path], Path],
) -> ScenarioResult:
    """Materialize one scenario, roundtrip it, classify the outcome."""
    started = time.monotonic()
    try:
        docx_path = scenario_dir / f"{spec.id}.docx"
        materialize(spec, matrix, docx_path)
        result = roundtrip(docx_path, timeout=60.0)
        input_children = extract_first(
            str(result.input_path), spec.parent_local, WORDPROCESSINGML
        )
        output_children = extract_first(
            str(result.output_path), spec.parent_local, WORDPROCESSINGML
        )
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
    return {
        "total": len(results),
        "preserved": sum(1 for r in results if r.verdict == "preserved"),
        "repaired": sum(1 for r in results if r.verdict == "repaired"),
        "missing": sum(1 for r in results if r.verdict == "missing"),
        "engine_error": sum(1 for r in results if r.verdict == "engine_error"),
        "any_repair_dialog": any(r.repair_dialog_seen for r in results),
        "repair_dialogs_count": sum(1 for r in results if r.repair_dialog_seen),
    }


def run_matrix_oracle(matrix: OrderingMatrix, *, output_path: Path | None = None) -> dict:
    """Run the full scenario matrix for one constraint type and emit JSON."""
    scenarios = all_scenarios(matrix)
    work_root = (
        Path.home() / "Documents" / ".word_oracle_runs" /
        f"{matrix.parent_local}-batch-{uuid.uuid4().hex[:8]}"
    )
    work_root.mkdir(parents=True, exist_ok=False)

    print(
        f"Running {len(scenarios)} {matrix.parent_local} scenarios → {work_root}",
        file=sys.stderr,
    )

    results: list[ScenarioResult] = []
    for i, spec in enumerate(scenarios, 1):
        print(
            f"  [{i}/{len(scenarios)}] {spec.id} ...", file=sys.stderr, end="", flush=True
        )
        result = _run_scenario(spec, matrix, work_root, matrix.materialize)
        results.append(result)
        marker = {
            "preserved": "·", "repaired": "!", "missing": "?", "engine_error": "X"
        }[result.verdict]
        dialog = " (dialog)" if result.repair_dialog_seen else ""
        print(f" {marker}{dialog} {result.elapsed_seconds:.1f}s", file=sys.stderr)

    report = {
        "constraint": f"CT_{matrix.parent_local[0].upper() + matrix.parent_local[1:]}",
        "spec_section": matrix.spec_section,
        "canonical_children_tested": list(matrix.canonical_children),
        "engine": "spec 011 Phase 1 (AppleScript open + close-with-save)",
        "word_version": word_version(),
        "platform": f"{platform.system()} {platform.release()}",
        "run_at": datetime.now(UTC).isoformat(timespec="seconds"),
        "summary": _summarize(results),
        "scenarios": [asdict(r) for r in results],
    }

    target = output_path or (
        BASELINES_DIR / f"word_{matrix.parent_local.lower()}_pairwise.json"
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(report, indent=2) + "\n")
    print(f"\nWrote {target}", file=sys.stderr)
    print(f"Summary: {report['summary']}", file=sys.stderr)

    shutil.rmtree(work_root, ignore_errors=True)
    _close_word_quietly()
    return report


# Backwards-compat alias kept for clarity in any earlier integrations.
def run_trpr_oracle(*, output_path: Path | None = None) -> dict:
    return run_matrix_oracle(MATRICES["trpr"], output_path=output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "constraint",
        choices=tuple(MATRICES),
        help="Which constraint type to run the oracle for.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help=(
            "Output path for the JSON baseline (default: "
            "tools/oracle/baselines/word_<constraint>_pairwise.json)."
        ),
    )
    args = parser.parse_args()

    run_matrix_oracle(MATRICES[args.constraint], output_path=args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
