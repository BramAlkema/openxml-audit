"""ODF roundtrip oracle: does a file survive soffice unchanged?

For each input ODF file:

  1. Roundtrip through `soffice --convert-to <format>` via the harness in
     `odf_window.py` (crash-resistant; per-call ephemeral profile).
  2. Extract the input's and the output's `content.xml` (the canonical
     ODF body part).
  3. Compute a content fingerprint and a structural diff: did soffice
     preserve the body verbatim, or did it rewrite it?
  4. Emit a per-file `RoundtripObservation` recording the outcome,
     timings, and a categorical signal (preserved / repaired / crash /
     timeout / open_failed).

This is the ODF equivalent of the Word roundtrip oracle in
`word_repair_oracle.py`. Where Word's oracle observes the
"unreadable content" repair dialog, soffice's headless mode is silent
about repairs — so the oracle infers them from the byte-level diff
between input and re-saved output.

A `preserved` outcome means the input was already in soffice's
canonical form. `repaired` means soffice changed something; the diff
has to be inspected to determine whether the change is cosmetic
(formatting reflow, attribute order) or substantive (element added,
removed, or reordered). The oracle commits the diffs as evidence
under `tools/oracle/baselines/` so future runs can detect drift in
soffice's repair behavior.

Spec: 0.6.3 (Spec 019).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Literal

REPO_ROOT = Path(__file__).resolve().parents[2]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.odf_window import (  # noqa: E402  (path setup must happen first)
    SofficeNotFoundError,
    SofficeRunResult,
    find_soffice,
    roundtrip,
)

from openxml_audit.package_diff import compare_packages  # noqa: E402


_CANONICAL_PARTS = ("content.xml", "styles.xml", "meta.xml", "settings.xml")


def _is_canonical_odf_part(name: str) -> bool:
    return name in _CANONICAL_PARTS
_FORMAT_BY_EXT = {
    ".odt": "odt",
    ".ods": "ods",
    ".odp": "odp",
    ".odg": "odg",
}


@dataclass
class PartFingerprint:
    """Sha256 of a single ODF part as it sits in the package's ZIP."""

    name: str
    sha256: str
    size_bytes: int


@dataclass
class RoundtripObservation:
    """Per-file outcome of a soffice roundtrip."""

    source_relpath: str
    target_format: str
    outcome: Literal[
        "preserved",  # input parts and output parts are byte-identical
        "repaired",   # output parts differ from input (cosmetic or substantive)
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
    diff_dir: str | None = None
    soffice_outcome: str = ""
    notes: list[str] = field(default_factory=list)


def _fingerprint_parts(odf_path: Path) -> list[PartFingerprint]:
    """Hash the canonical ODF parts inside the package.

    Only fingerprints `content.xml` / `styles.xml` / `meta.xml` /
    `settings.xml` (when present). Skips manifest, mimetype, and binary
    media — those are the parts soffice rewrites unconditionally and
    they generate noise in the diff.
    """
    fingerprints: list[PartFingerprint] = []
    if not odf_path.exists():
        return fingerprints
    with zipfile.ZipFile(odf_path) as zf:
        names = set(zf.namelist())
        for name in _CANONICAL_PARTS:
            if name not in names:
                continue
            data = zf.read(name)
            fingerprints.append(PartFingerprint(
                name=name,
                sha256=hashlib.sha256(data).hexdigest(),
                size_bytes=len(data),
            ))
    return fingerprints


def _detect_format(input_path: Path) -> str:
    fmt = _FORMAT_BY_EXT.get(input_path.suffix.lower())
    if fmt is None:
        raise ValueError(
            f"unsupported ODF extension '{input_path.suffix}' for {input_path}"
        )
    return fmt


def observe(
    input_path: Path,
    work_dir: Path,
    *,
    timeout_seconds: float = 60.0,
    keep_artifacts: bool = False,
) -> RoundtripObservation:
    """Roundtrip one file and observe the outcome.

    `work_dir` is used for the soffice output. The caller is responsible
    for choosing a per-file directory if avoiding name collisions in
    batch mode. When `keep_artifacts=True` the per-part diff directory
    under `work_dir/compare/` is preserved so the caller can inspect
    what soffice actually changed.
    """
    target_format = _detect_format(input_path)
    input_parts = _fingerprint_parts(input_path)

    run: SofficeRunResult = roundtrip(
        input_path,
        work_dir,
        target_format=target_format,  # type: ignore[arg-type]
        timeout_seconds=timeout_seconds,
    )

    if run.outcome == "timeout":
        return RoundtripObservation(
            source_relpath=str(input_path.name),
            target_format=target_format,
            outcome="timeout",
            duration_seconds=run.duration_seconds,
            input_parts=input_parts,
            soffice_outcome=run.outcome,
            notes=run.notes,
        )
    if run.outcome == "exit_nonzero":
        # soffice's headless converter usually exits nonzero only when
        # it fails to open the input. Treat as open_failed.
        return RoundtripObservation(
            source_relpath=str(input_path.name),
            target_format=target_format,
            outcome="open_failed",
            duration_seconds=run.duration_seconds,
            input_parts=input_parts,
            soffice_outcome=run.outcome,
            notes=run.notes + [f"return_code={run.return_code}", run.stderr.strip()[:500]],
        )
    if run.outcome == "missing_output" or run.output_path is None:
        return RoundtripObservation(
            source_relpath=str(input_path.name),
            target_format=target_format,
            outcome="missing_output",
            duration_seconds=run.duration_seconds,
            input_parts=input_parts,
            soffice_outcome=run.outcome,
            notes=run.notes,
        )
    if run.outcome != "ok":
        return RoundtripObservation(
            source_relpath=str(input_path.name),
            target_format=target_format,
            outcome="crash",
            duration_seconds=run.duration_seconds,
            input_parts=input_parts,
            soffice_outcome=run.outcome,
            notes=run.notes,
        )

    output_parts = _fingerprint_parts(run.output_path)

    # Per-part canonical-c14n diff via the shared package_diff module.
    # Replaces hash-only diff used through 0.6.7; the report includes
    # per-part text diffs under work_dir/compare/diffs/ so callers
    # can inspect what soffice actually changed (cosmetic XML reflow
    # vs substantive content edit).
    compare_dir = work_dir / "compare"
    diff_report = compare_packages(
        base_path=input_path,
        head_path=run.output_path,
        output_dir=compare_dir,
        parts_filter=_is_canonical_odf_part,
    )
    changed = list(diff_report.get("changed_files", []))
    added = list(diff_report.get("added_files", []))
    removed = list(diff_report.get("removed_files", []))

    return RoundtripObservation(
        source_relpath=str(input_path.name),
        target_format=target_format,
        outcome="preserved" if not (changed or added or removed) else "repaired",
        duration_seconds=run.duration_seconds,
        input_parts=input_parts,
        output_parts=output_parts,
        changed_parts=changed,
        added_parts=added,
        removed_parts=removed,
        diff_dir=str(compare_dir) if keep_artifacts else None,
        soffice_outcome=run.outcome,
        notes=run.notes,
    )


def observe_batch(
    inputs: list[Path],
    work_root: Path,
    *,
    timeout_seconds: float = 60.0,
    keep_artifacts: bool = False,
) -> list[RoundtripObservation]:
    """Run `observe` over a batch, allocating per-file work dirs."""
    observations: list[RoundtripObservation] = []
    for index, input_path in enumerate(inputs):
        work_dir = work_root / f"{index:04d}-{input_path.stem}"
        try:
            obs = observe(
                input_path, work_dir,
                timeout_seconds=timeout_seconds,
                keep_artifacts=keep_artifacts,
            )
        except ValueError as exc:
            obs = RoundtripObservation(
                source_relpath=str(input_path.name),
                target_format=input_path.suffix.lower().lstrip("."),
                outcome="open_failed",
                duration_seconds=0.0,
                notes=[f"unsupported input: {exc}"],
            )
        observations.append(obs)
    return observations


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
        },
    }


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("input", nargs="+", type=Path,
                   help="ODF files to roundtrip (or directories to walk)")
    p.add_argument("--work-root", type=Path, default=Path("/tmp/odf_oracle_runs"),
                   help="root directory for soffice output (default /tmp/...)")
    p.add_argument("--output", type=Path, default=None,
                   help="optional path to write the JSON observation report")
    p.add_argument("--timeout", type=float, default=60.0,
                   help="per-file soffice timeout in seconds (default 60)")
    p.add_argument("--keep-artifacts", action="store_true",
                   help="leave per-file work dirs in place for inspection")
    args = p.parse_args()

    try:
        find_soffice()
    except SofficeNotFoundError as exc:
        print(f"soffice unavailable: {exc}", file=sys.stderr)
        return 2

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(sorted(
                p for p in entry.rglob("*")
                if p.is_file() and p.suffix.lower() in _FORMAT_BY_EXT
            ))
        elif entry.is_file():
            inputs.append(entry)

    if not inputs:
        print("no ODF inputs found", file=sys.stderr)
        return 2

    args.work_root.mkdir(parents=True, exist_ok=True)
    observations = observe_batch(
        inputs, args.work_root,
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
        f"\nodf-oracle: total={summary['total']} "
        f"preserved={summary['preserved']} repaired={summary['repaired']} "
        f"crash={summary['crash']} timeout={summary['timeout']} "
        f"open_failed={summary['open_failed']}",
        file=sys.stderr,
    )
    # Exit nonzero if anything errored hard, but not on `repaired` —
    # that's data, not a failure.
    hard_errors = summary["crash"] + summary["timeout"] + summary["open_failed"] + summary["missing_output"]
    return 1 if hard_errors > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
