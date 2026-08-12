"""Excel roundtrip oracle: does a workbook survive Excel unchanged?

Phase 1 (0.6.5) of Spec 021. Fourth in the family:

  - `tools/oracle/word_repair_oracle.py`  — Word for Mac (Spec 011)
  - `tools/oracle/odf_repair_oracle.py`   — LibreOffice (Spec 019)
  - `tools/oracle/pptx_repair_oracle.py`  — PowerPoint (Spec 020)
  - `tools/oracle/xlsx_repair_oracle.py`  — this file (Spec 021)

Built on top of the existing `openxml_audit.xlsx.osa` layer; this
release extends that module with the missing primitives
(`close_workbook_saving`, `is_workbook_open`,
`find_repair_dialog_text`, `REPAIR_DIALOG_PATTERNS`,
`list_open_workbook_names`) — the Word/PowerPoint siblings already
have these.

Excel does not yet have a `xlsx.lab` package differ analogous to
`pptx.lab.compare_pptx_packages`. For Phase 1 we use a hash-based
per-part diff (same approach the ODF oracle uses), fingerprinting
the canonical OOXML parts of an XLSX/XLTX/XLSM:

  - [Content_Types].xml
  - xl/workbook.xml
  - xl/_rels/workbook.xml.rels
  - xl/worksheets/*.xml
  - xl/styles.xml
  - xl/sharedStrings.xml
  - xl/theme/*.xml
  - xl/charts/*.xml

Per file:

  1. Stage the input under `~/Documents/.xlsx_oracle_runs/<id>/`.
     Excel's App Sandbox grants access there by default; tmpfs
     paths are blocked (override with `XLSX_ORACLE_STAGE`).
  2. Launch Excel, open the staged copy, poll until it registers
     in Excel's `workbooks` collection.
  3. Detect (don't auto-dismiss) Excel's "found a problem" modal
     if it appears. Phase 1 records dialog text as evidence;
     auto-dismiss is Phase 2.
  4. Close-with-save. Excel writes the modified workbook back
     to the staged path before tearing down the window.
  5. Fingerprint the canonical parts before/after; outcome is
     `preserved` (no parts changed), `repaired` (>=1 changed/added/
     removed), or one of `crash` / `timeout` / `open_failed`.

Spec: `specs/021-xlsx-roundtrip-oracle.md`.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Literal

from openxml_audit.package_diff import compare_packages
from openxml_audit.xlsx import osa as xlsx_osa

_DEFAULT_STAGE_PARENT = Path.home() / "Documents" / ".xlsx_oracle_runs"

# OOXML parts we fingerprint for the oracle's diff. Skips media
# (binary) and printerSettings (regenerated on save).
_FINGERPRINTED_PREFIXES = (
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/theme/",
    "xl/worksheets/",
    "xl/chartsheets/",
    "xl/charts/",
    "xl/tables/",
    "xl/pivotTables/",
    "xl/pivotCache/",
    "xl/queryTables/",
    "xl/connections.xml",
    "xl/calcChain.xml",
    "xl/externalLinks/",
)

_SUPPORTED_EXTS = {".xlsx", ".xltx", ".xlsm", ".xltm", ".xlsb"}


def _is_fingerprinted_part(name: str) -> bool:
    return any(
        name == prefix or (prefix.endswith("/") and name.startswith(prefix))
        for prefix in _FINGERPRINTED_PREFIXES
    )


@dataclass
class PartFingerprint:
    name: str
    sha256: str
    size_bytes: int


@dataclass
class RoundtripObservation:
    """Per-file outcome of an Excel roundtrip."""

    source_relpath: str
    outcome: Literal[
        "preserved",
        "repaired",
        "crash",
        "timeout",
        "open_failed",
        "missing_output",
    ]
    duration_seconds: float
    input_parts: list[PartFingerprint] = field(default_factory=list)
    output_parts: list[PartFingerprint] = field(default_factory=list)
    changed_parts: list[str] = field(default_factory=list)
    added_parts: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    diff_dir: str | None = None  # path to per-part diff artifacts
    repair_dialog_seen: bool = False
    repair_dialog_text: str | None = None
    repair_dialog_button_clicked: str | None = None
    notes: list[str] = field(default_factory=list)


def _stage_root() -> Path:
    """Resolve the staging root, defaulting to ~/Documents/.xlsx_oracle_runs.

    Excel for Mac's App Sandbox grants access to ~/Documents by
    default; /tmp and other locations require explicit Full Disk
    Access. Override with `XLSX_ORACLE_STAGE`.
    """
    override = os.environ.get("XLSX_ORACLE_STAGE")
    base = Path(override).expanduser() if override else _DEFAULT_STAGE_PARENT
    base.mkdir(parents=True, exist_ok=True)
    return base


def _stage_input(input_path: Path) -> tuple[Path, Path]:
    """Stage the input under a per-run work dir.

    Returns (work_dir, working_copy). The working copy is what Excel
    will overwrite when it closes-with-save.
    """
    work_dir = _stage_root() / f"run-{uuid.uuid4().hex[:8]}"
    work_dir.mkdir(parents=True, exist_ok=False)
    working_copy = work_dir / input_path.name
    shutil.copy2(input_path, working_copy)
    return work_dir, working_copy


def _fingerprint_xlsx(xlsx_path: Path) -> list[PartFingerprint]:
    fingerprints: list[PartFingerprint] = []
    if not xlsx_path.exists():
        return fingerprints
    with zipfile.ZipFile(xlsx_path) as zf:
        for name in sorted(zf.namelist()):
            if not _is_fingerprinted_part(name):
                continue
            data = zf.read(name)
            fingerprints.append(
                PartFingerprint(
                    name=name,
                    sha256=hashlib.sha256(data).hexdigest(),
                    size_bytes=len(data),
                )
            )
    return fingerprints


def _wait_for_open(staged: Path, timeout: float, *, poll: float = 0.5) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if xlsx_osa.is_workbook_open(staged):
            return True
        time.sleep(poll)
    return False


def observe(
    input_xlsx: Path,
    *,
    timeout_seconds: float = 60.0,
    keep_artifacts: bool = False,
) -> RoundtripObservation:
    """Roundtrip one workbook through Excel and record the outcome."""
    input_xlsx = input_xlsx.resolve()
    if not input_xlsx.exists():
        return RoundtripObservation(
            source_relpath=str(input_xlsx.name),
            outcome="missing_output",
            duration_seconds=0.0,
            notes=[f"input does not exist: {input_xlsx}"],
        )

    input_parts = _fingerprint_xlsx(input_xlsx)
    work_dir, staged = _stage_input(input_xlsx)
    started = time.monotonic()
    repair_seen = False
    repair_text: str | None = None
    repair_button_clicked: str | None = None

    try:
        xlsx_osa.launch_excel()
        time.sleep(1.0)

        try:
            xlsx_osa.open_workbook(staged, timeout=10.0)
        except (RuntimeError, subprocess.TimeoutExpired) as exc:
            return RoundtripObservation(
                source_relpath=str(input_xlsx.name),
                outcome="open_failed",
                duration_seconds=time.monotonic() - started,
                input_parts=input_parts,
                notes=[f"open_workbook failed: {exc}"],
            )

        # Interleave repair-dialog auto-dismiss with the open-poll: the
        # dialog can block the workbook from registering as open.
        deadline = time.monotonic() + timeout_seconds
        seen_open = False
        while time.monotonic() < deadline:
            if not repair_seen:
                seen, text, clicked = xlsx_osa.dismiss_repair_dialog()
                if seen:
                    repair_seen = True
                    repair_text = text
                    repair_button_clicked = clicked
                    if clicked is None:
                        return RoundtripObservation(
                            source_relpath=str(input_xlsx.name),
                            outcome="open_failed",
                            duration_seconds=time.monotonic() - started,
                            input_parts=input_parts,
                            repair_dialog_seen=True,
                            repair_dialog_text=text,
                            repair_dialog_button_clicked=None,
                            notes=[
                                "Repair dialog detected but no accept "
                                "button label matched — manual dismissal "
                                "required, oracle bailed."
                            ],
                        )
            if xlsx_osa.is_workbook_open(staged):
                seen_open = True
                break
            time.sleep(0.5)

        if not seen_open:
            return RoundtripObservation(
                source_relpath=str(input_xlsx.name),
                outcome="open_failed",
                duration_seconds=time.monotonic() - started,
                input_parts=input_parts,
                repair_dialog_seen=repair_seen,
                repair_dialog_text=repair_text,
                repair_dialog_button_clicked=repair_button_clicked,
                notes=[
                    f"Excel did not register {staged.name!r} as open within {timeout_seconds:.0f}s"
                ],
            )

        # Follow-up dialog scan: Excel sometimes shows a "repairs were
        # made" info modal after the primary recovery prompt.
        if not repair_seen:
            seen, text, clicked = xlsx_osa.dismiss_repair_dialog()
            if seen:
                repair_seen = True
                repair_text = text
                repair_button_clicked = clicked

        try:
            xlsx_osa.close_workbook_saving()
        except (RuntimeError, subprocess.TimeoutExpired) as exc:
            return RoundtripObservation(
                source_relpath=str(input_xlsx.name),
                outcome="crash",
                duration_seconds=time.monotonic() - started,
                input_parts=input_parts,
                repair_dialog_seen=repair_seen,
                repair_dialog_text=repair_text,
                repair_dialog_button_clicked=repair_button_clicked,
                notes=[f"close_workbook_saving failed: {exc}"],
            )

        # After close-with-save, Excel may leave a "we made repairs"
        # info modal up (View/Delete buttons that aren't in our accept
        # list). Send Escape to dismiss any leftover modal so the next
        # corpus item starts on a clean Excel window.
        xlsx_osa.dismiss_any_leftover_modal()

        if not staged.exists():
            return RoundtripObservation(
                source_relpath=str(input_xlsx.name),
                outcome="missing_output",
                duration_seconds=time.monotonic() - started,
                input_parts=input_parts,
                repair_dialog_seen=repair_seen,
                repair_dialog_text=repair_text,
                repair_dialog_button_clicked=repair_button_clicked,
                notes=["close-with-save returned but staged file is missing"],
            )

        output_parts = _fingerprint_xlsx(staged)

        # Per-part canonical-c14n diff via the shared package_diff
        # module. Replaces the hash-only diff used through 0.6.7;
        # the report includes per-part text diffs under
        # work_dir/compare/diffs/ so callers can inspect what
        # Excel actually changed (rather than just "something
        # changed in this file").
        compare_dir = work_dir / "compare"
        report = compare_packages(
            base_path=input_xlsx,
            head_path=staged,
            output_dir=compare_dir,
            parts_filter=_is_fingerprinted_part,
        )
        changed = list(report.get("changed_files", []))
        added = list(report.get("added_files", []))
        removed = list(report.get("removed_files", []))
        outcome: Literal["preserved", "repaired"] = (
            "preserved" if not (changed or added or removed) else "repaired"
        )

        return RoundtripObservation(
            source_relpath=str(input_xlsx.name),
            outcome=outcome,
            duration_seconds=time.monotonic() - started,
            input_parts=input_parts,
            output_parts=output_parts,
            changed_parts=changed,
            added_parts=added,
            removed_parts=removed,
            diff_dir=str(compare_dir) if keep_artifacts else None,
            repair_dialog_seen=repair_seen,
            repair_dialog_text=repair_text,
            repair_dialog_button_clicked=repair_button_clicked,
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
        observe(p, timeout_seconds=timeout_seconds, keep_artifacts=keep_artifacts) for p in inputs
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
    p.add_argument(
        "input",
        nargs="+",
        type=Path,
        help=".xlsx/.xltx/.xlsm files to roundtrip (or directories to walk)",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="optional path to write the JSON observation report",
    )
    p.add_argument(
        "--timeout", type=float, default=60.0, help="per-file Excel timeout in seconds (default 60)"
    )
    p.add_argument(
        "--keep-artifacts", action="store_true", help="leave staging dirs in place for inspection"
    )
    args = p.parse_args()

    if not xlsx_osa.EXCEL_APP_BUNDLE.exists():
        print(
            f"Microsoft Excel not installed at {xlsx_osa.EXCEL_APP_BUNDLE}",
            file=sys.stderr,
        )
        return 2

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(
                sorted(
                    p
                    for p in entry.rglob("*")
                    if p.is_file() and p.suffix.lower() in _SUPPORTED_EXTS
                )
            )
        elif entry.is_file() and entry.suffix.lower() in _SUPPORTED_EXTS:
            inputs.append(entry)

    if not inputs:
        print("no Excel inputs found", file=sys.stderr)
        return 2

    observations = observe_batch(
        inputs,
        timeout_seconds=args.timeout,
        keep_artifacts=args.keep_artifacts,
    )
    report = _to_jsonable(observations)

    if args.output:
        args.output.write_text(json.dumps(report, indent=2))
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(json.dumps(report, indent=2))

    summary = report["summary"]
    print(
        f"\nxlsx-oracle: total={summary['total']} "
        f"preserved={summary['preserved']} repaired={summary['repaired']} "
        f"crash={summary['crash']} timeout={summary['timeout']} "
        f"open_failed={summary['open_failed']} "
        f"repair_dialog_seen={summary['repair_dialog_seen']}",
        file=sys.stderr,
    )
    hard = (
        summary["crash"] + summary["timeout"] + summary["open_failed"] + summary["missing_output"]
    )
    return 1 if hard > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
