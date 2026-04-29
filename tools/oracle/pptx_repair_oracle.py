"""PowerPoint roundtrip oracle: does a file survive PowerPoint unchanged?

Phase 1 (0.6.4) of Spec 020. Sibling of:

  - `tools/oracle/word_repair_oracle.py` — Word for Mac via osascript
  - `tools/oracle/odf_repair_oracle.py`  — LibreOffice via headless soffice

This orchestrator deliberately does **not** duplicate primitives that
already live in `openxml_audit.pptx.osa` (window control, save, close,
repair-dialog detection) or in `openxml_audit.pptx.lab`
(`compare_pptx_packages` — full per-part snapshot/diff infrastructure).
It wires those existing layers into the standard observation shape.

Per file:

  1. Stage the input under `~/Documents/.pptx_oracle_runs/<id>/`.
     PowerPoint's App Sandbox grants access there by default; tmpfs
     paths are blocked.
  2. Launch PowerPoint, open the staged copy, poll until it registers
     in PowerPoint's `presentations` collection.
  3. Detect (don't auto-dismiss) PowerPoint's repair dialog if it
     appears — Phase 1 records dialog text as evidence; Phase 2 adds
     button-click ceremony like the Word oracle's.
  4. Close-with-save. PowerPoint writes the modified presentation back
     to the staged path before tearing down the window.
  5. Hand the original + the post-PowerPoint staged copy to
     `pptx.lab.compare_pptx_packages`, which writes per-part diffs
     under a work directory and returns a structured report.
  6. Roll up the report into a `RoundtripObservation`.

Spec: `specs/020-pptx-roundtrip-oracle.md`.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Literal

from openxml_audit.pptx import osa as pptx_osa
from openxml_audit.pptx.lab import compare_pptx_packages


_DEFAULT_STAGE_PARENT = Path.home() / "Documents" / ".pptx_oracle_runs"


@dataclass
class RoundtripObservation:
    """Per-file outcome of a PowerPoint roundtrip."""

    source_relpath: str
    outcome: Literal[
        "preserved",  # 0 changed/added/removed parts in the package diff
        "repaired",   # >=1 changed/added/removed parts
        "crash",
        "timeout",
        "open_failed",
        "missing_output",
    ]
    duration_seconds: float
    changed_parts: list[str] = field(default_factory=list)
    added_parts: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    diff_dir: str | None = None  # path to per-part diff artifacts
    repair_dialog_seen: bool = False
    repair_dialog_text: str | None = None
    notes: list[str] = field(default_factory=list)


def _stage_root() -> Path:
    """Resolve the staging root, defaulting to ~/Documents/.pptx_oracle_runs.

    PowerPoint for Mac's App Sandbox grants access to ~/Documents by
    default; /tmp and other locations require explicit Full Disk
    Access. Override with `PPTX_ORACLE_STAGE`.
    """
    override = os.environ.get("PPTX_ORACLE_STAGE")
    base = Path(override).expanduser() if override else _DEFAULT_STAGE_PARENT
    base.mkdir(parents=True, exist_ok=True)
    return base


def _stage_input(input_path: Path) -> tuple[Path, Path, Path]:
    """Stage the input under a per-run work dir.

    Returns (work_dir, original_copy, working_copy). `original_copy`
    is preserved verbatim for the diff; `working_copy` is what
    PowerPoint will overwrite when it closes-with-save.
    """
    work_dir = _stage_root() / f"run-{uuid.uuid4().hex[:8]}"
    work_dir.mkdir(parents=True, exist_ok=False)
    original_copy = work_dir / f"original_{input_path.name}"
    working_copy = work_dir / input_path.name
    shutil.copy2(input_path, original_copy)
    shutil.copy2(input_path, working_copy)
    return work_dir, original_copy, working_copy


def _wait_for_open(staged: Path, timeout: float, *, poll: float = 0.5) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if pptx_osa.is_presentation_open(staged):
            return True
        time.sleep(poll)
    return False


def observe(
    input_pptx: Path,
    *,
    timeout_seconds: float = 60.0,
    keep_artifacts: bool = False,
) -> RoundtripObservation:
    """Roundtrip one .pptx through PowerPoint and record the outcome.

    `keep_artifacts=True` leaves the staging directory in place so the
    caller can inspect the per-part diffs under `<work_dir>/compare/`.
    """
    input_pptx = input_pptx.resolve()
    if not input_pptx.exists():
        return RoundtripObservation(
            source_relpath=str(input_pptx.name),
            outcome="missing_output",
            duration_seconds=0.0,
            notes=[f"input does not exist: {input_pptx}"],
        )

    work_dir, original, working = _stage_input(input_pptx)
    started = time.monotonic()
    repair_seen = False
    repair_text: str | None = None

    try:
        pptx_osa.launch_powerpoint()
        # Settling pause — `open` doesn't block until PowerPoint is
        # AppleScript-reachable. The first osascript against a
        # cold-launched PowerPoint can race the dispatcher.
        time.sleep(1.0)

        try:
            pptx_osa.open_presentation_via_ui(working, timeout=10.0)
        except (RuntimeError, subprocess.TimeoutExpired) as exc:
            return RoundtripObservation(
                source_relpath=str(input_pptx.name),
                outcome="open_failed",
                duration_seconds=time.monotonic() - started,
                notes=[f"open_presentation_via_ui failed: {exc}"],
            )

        if not _wait_for_open(working, timeout=timeout_seconds):
            text = pptx_osa.find_repair_dialog_text()
            if text is not None:
                repair_seen = True
                repair_text = text
            return RoundtripObservation(
                source_relpath=str(input_pptx.name),
                outcome="open_failed",
                duration_seconds=time.monotonic() - started,
                repair_dialog_seen=repair_seen,
                repair_dialog_text=repair_text,
                notes=[
                    f"PowerPoint did not register {working.name!r} as open "
                    f"within {timeout_seconds:.0f}s"
                ],
            )

        # Once open, scan for a repair dialog (modal next to the open
        # presentation; PowerPoint may surface it asynchronously).
        text = pptx_osa.find_repair_dialog_text()
        if text is not None:
            repair_seen = True
            repair_text = text

        try:
            pptx_osa.close_presentation_saving()
        except (RuntimeError, subprocess.TimeoutExpired) as exc:
            return RoundtripObservation(
                source_relpath=str(input_pptx.name),
                outcome="crash",
                duration_seconds=time.monotonic() - started,
                repair_dialog_seen=repair_seen,
                repair_dialog_text=repair_text,
                notes=[f"close_presentation_saving failed: {exc}"],
            )

        if not working.exists():
            return RoundtripObservation(
                source_relpath=str(input_pptx.name),
                outcome="missing_output",
                duration_seconds=time.monotonic() - started,
                repair_dialog_seen=repair_seen,
                repair_dialog_text=repair_text,
                notes=["close-with-save returned but staged file is missing"],
            )

        # Hand off to pptx.lab for the per-part diff.
        compare_dir = work_dir / "compare"
        report = compare_pptx_packages(
            base_path=original,
            head_path=working,
            output_dir=compare_dir,
        )
        changed = list(report.get("changed_files", []))
        added = list(report.get("added_files", []))
        removed = list(report.get("removed_files", []))
        outcome: Literal["preserved", "repaired"] = (
            "preserved" if not (changed or added or removed) else "repaired"
        )

        return RoundtripObservation(
            source_relpath=str(input_pptx.name),
            outcome=outcome,
            duration_seconds=time.monotonic() - started,
            changed_parts=changed,
            added_parts=added,
            removed_parts=removed,
            diff_dir=str(compare_dir) if keep_artifacts else None,
            repair_dialog_seen=repair_seen,
            repair_dialog_text=repair_text,
        )
    finally:
        if not keep_artifacts:
            shutil.rmtree(work_dir, ignore_errors=True)


def observe_batch(
    inputs: list[Path],
    *,
    timeout_seconds: float = 60.0,
    keep_artifacts: bool = False,
) -> list[RoundtripObservation]:
    return [
        observe(p, timeout_seconds=timeout_seconds, keep_artifacts=keep_artifacts)
        for p in inputs
    ]


def _to_jsonable(observations: list[RoundtripObservation]) -> dict:
    return {
        "schema_version": 1,
        "observations": [asdict(obs) for obs in observations],
        "summary": {
            "total": len(observations),
            "preserved": sum(1 for o in observations if o.outcome == "preserved"),
            "repaired": sum(1 for o in observations if o.outcome == "repaired"),
            "crash": sum(1 for o in observations if o.outcome == "crash"),
            "timeout": sum(1 for o in observations if o.outcome == "timeout"),
            "open_failed": sum(1 for o in observations if o.outcome == "open_failed"),
            "missing_output": sum(1 for o in observations if o.outcome == "missing_output"),
            "repair_dialog_seen": sum(1 for o in observations if o.repair_dialog_seen),
        },
    }


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("input", nargs="+", type=Path,
                   help=".pptx files to roundtrip (or directories to walk)")
    p.add_argument("--output", type=Path, default=None,
                   help="optional path to write the JSON observation report")
    p.add_argument("--timeout", type=float, default=60.0,
                   help="per-file PowerPoint timeout in seconds (default 60)")
    p.add_argument("--keep-artifacts", action="store_true",
                   help="leave staging dirs in place for inspection")
    args = p.parse_args()

    if not pptx_osa.POWERPOINT_APP_BUNDLE.exists():
        print(
            "Microsoft PowerPoint not installed at "
            f"{pptx_osa.POWERPOINT_APP_BUNDLE}",
            file=sys.stderr,
        )
        return 2

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(sorted(entry.rglob("*.pptx")))
        elif entry.is_file() and entry.suffix.lower() == ".pptx":
            inputs.append(entry)

    if not inputs:
        print("no .pptx inputs found", file=sys.stderr)
        return 2

    observations = observe_batch(
        inputs, timeout_seconds=args.timeout, keep_artifacts=args.keep_artifacts,
    )
    report = _to_jsonable(observations)

    if args.output:
        args.output.write_text(json.dumps(report, indent=2))
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(json.dumps(report, indent=2))

    summary = report["summary"]
    print(
        f"\npptx-oracle: total={summary['total']} "
        f"preserved={summary['preserved']} repaired={summary['repaired']} "
        f"crash={summary['crash']} timeout={summary['timeout']} "
        f"open_failed={summary['open_failed']} "
        f"repair_dialog_seen={summary['repair_dialog_seen']}",
        file=sys.stderr,
    )
    hard = summary["crash"] + summary["timeout"] + summary["open_failed"] + summary["missing_output"]
    return 1 if hard > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
