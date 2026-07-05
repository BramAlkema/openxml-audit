"""LibreOffice OOXML roundtrip oracle: the CI-runnable rung (Spec 036).

For each input `.docx` / `.xlsx` / `.pptx`:

  1. Roundtrip through headless soffice (`--convert-to <same format>`)
     via the supervised harness in `odf_window.py`.
  2. Fingerprint and diff the format's canonical parts (shared
     `package_diff` module), recording changed/added/removed parts.
  3. Emit a per-file `RoundtripObservation` (same schema as the ODF
     oracle: preserved / repaired / crash / timeout / open_failed).

Expectation calibration: LibreOffice virtually always rewrites foreign
formats at the byte level, so `preserved` is rare by construction.
The mission-relevant signals are:

  - does LibreOffice open the file at all (open_failed/crash)?
  - does it drop or add parts (slides disappearing is substantive)?
  - with `--reference-manifest`: which reference-document features
    survive the roundtrip structurally (feature_probes signatures)?

Observations are evidence about LibreOffice as a target app. They say
nothing about Word/Excel/PowerPoint and do not feed registry tier
promotion (tiers are defined against the file's own target app).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Literal

REPO_ROOT = Path(__file__).resolve().parents[2]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from openxml_audit.package_diff import compare_packages  # noqa: E402
from openxml_audit.reference.feature_probes import (  # noqa: E402
    PPTX_FEATURE_SIGNATURES,
    probe_slide,
)
from oracle.odf_window import (  # noqa: E402  (path setup must happen first)
    SofficeNotFoundError,
    SofficeRunResult,
    find_soffice,
    roundtrip,
)

_FORMAT_BY_EXT = {
    ".docx": "docx",
    ".xlsx": "xlsx",
    ".pptx": "pptx",
}

# Canonical parts per format: the evidence-bearing XML the oracle
# fingerprints and diffs. Media, rels, docProps, and printer settings
# are rewritten unconditionally by converters and only add noise.
_CANONICAL_PREFIXES = {
    "docx": ("word/document.xml", "word/styles.xml", "word/numbering.xml"),
    "xlsx": (
        "xl/workbook.xml",
        "xl/worksheets/",
        "xl/sharedStrings.xml",
        "xl/styles.xml",
    ),
    "pptx": ("ppt/presentation.xml", "ppt/slides/"),
}

_SLIDE_LOCATION = re.compile(r"^slides? (\d+)(?:-(\d+))?$")


def _canonical_filter(fmt: str):
    prefixes = _CANONICAL_PREFIXES[fmt]

    def _is_canonical(name: str) -> bool:
        return name.endswith(".xml") and any(
            name == prefix or (prefix.endswith("/") and name.startswith(prefix))
            for prefix in prefixes
        )

    return _is_canonical


@dataclass
class PartFingerprint:
    """Sha256 of a single canonical part as it sits in the package."""

    name: str
    sha256: str
    size_bytes: int


@dataclass
class FeatureSurvival:
    """Per-feature structural survival after the roundtrip."""

    key: str
    location: str
    slide_parts_present: bool
    signature_present: bool | None  # None: no signature registered


@dataclass
class RoundtripObservation:
    """Per-file outcome of a soffice OOXML roundtrip."""

    source_relpath: str
    target_format: str
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
    feature_survival: list[FeatureSurvival] = field(default_factory=list)
    diff_dir: str | None = None
    soffice_outcome: str = ""
    notes: list[str] = field(default_factory=list)


def _fingerprint_parts(package_path: Path, fmt: str) -> list[PartFingerprint]:
    fingerprints: list[PartFingerprint] = []
    if not package_path.exists():
        return fingerprints
    is_canonical = _canonical_filter(fmt)
    with zipfile.ZipFile(package_path) as zf:
        for name in sorted(zf.namelist()):
            if not is_canonical(name):
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


def _detect_format(input_path: Path) -> str:
    fmt = _FORMAT_BY_EXT.get(input_path.suffix.lower())
    if fmt is None:
        raise ValueError(f"unsupported OOXML extension '{input_path.suffix}' for {input_path}")
    return fmt


def _slide_parts_for_location(location: str) -> list[str]:
    """Map a manifest location ('slide 2', 'slides 7-8') to part names."""
    match = _SLIDE_LOCATION.match(location)
    if not match:
        return []
    first = int(match.group(1))
    last = int(match.group(2)) if match.group(2) else first
    return [f"ppt/slides/slide{number}.xml" for number in range(first, last + 1)]


def survey_feature_survival(
    output_path: Path,
    manifest: dict,
) -> list[FeatureSurvival]:
    """Probe a roundtripped PPTX for each manifest feature's survival."""
    survivals: list[FeatureSurvival] = []
    with zipfile.ZipFile(output_path) as zf:
        names = set(zf.namelist())
        for feature in manifest.get("features", []):
            key = feature.get("key", "")
            location = feature.get("location", "")
            slide_parts = _slide_parts_for_location(location)
            present = bool(slide_parts) and all(part in names for part in slide_parts)
            signature: bool | None = None
            if present and key in PPTX_FEATURE_SIGNATURES:
                signature = any(probe_slide(zf.read(part), key) for part in slide_parts)
            survivals.append(
                FeatureSurvival(
                    key=key,
                    location=location,
                    slide_parts_present=present,
                    signature_present=signature,
                )
            )
    return survivals


def observe(
    input_path: Path,
    work_dir: Path,
    *,
    timeout_seconds: float = 120.0,
    keep_artifacts: bool = False,
    reference_manifest: dict | None = None,
) -> RoundtripObservation:
    """Roundtrip one OOXML file through soffice and observe the outcome."""
    target_format = _detect_format(input_path)
    input_parts = _fingerprint_parts(input_path, target_format)

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

    output_parts = _fingerprint_parts(run.output_path, target_format)

    compare_dir = work_dir / "compare"
    diff_report = compare_packages(
        base_path=input_path,
        head_path=run.output_path,
        output_dir=compare_dir,
        parts_filter=_canonical_filter(target_format),
    )
    changed = list(diff_report.get("changed_files", []))
    added = list(diff_report.get("added_files", []))
    removed = list(diff_report.get("removed_files", []))

    survivals: list[FeatureSurvival] = []
    if reference_manifest is not None and target_format == "pptx":
        survivals = survey_feature_survival(run.output_path, reference_manifest)

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
        feature_survival=survivals,
        diff_dir=str(compare_dir) if keep_artifacts else None,
        soffice_outcome=run.outcome,
        notes=run.notes,
    )


def observe_batch(
    inputs: list[Path],
    work_root: Path,
    *,
    timeout_seconds: float = 120.0,
    keep_artifacts: bool = False,
    reference_manifests: dict[str, dict] | None = None,
) -> list[RoundtripObservation]:
    """Run `observe` over a batch, allocating per-file work dirs.

    `reference_manifests` maps input basenames to manifest payloads for
    feature-survival probing.
    """
    observations: list[RoundtripObservation] = []
    for index, input_path in enumerate(inputs):
        work_dir = work_root / f"{index:04d}-{input_path.stem}"
        manifest = (reference_manifests or {}).get(input_path.name)
        try:
            obs = observe(
                input_path,
                work_dir,
                timeout_seconds=timeout_seconds,
                keep_artifacts=keep_artifacts,
                reference_manifest=manifest,
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
        "oracle": "libreoffice-ooxml-roundtrip",
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
    p.add_argument(
        "input",
        nargs="+",
        type=Path,
        help="OOXML files to roundtrip (or directories to walk)",
    )
    p.add_argument(
        "--work-root",
        type=Path,
        default=Path("/tmp/lo_ooxml_oracle_runs"),
        help="root directory for soffice output (default /tmp/...)",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="optional path to write the JSON observation report",
    )
    p.add_argument(
        "--timeout",
        type=float,
        default=120.0,
        help="per-file soffice timeout in seconds (default 120)",
    )
    p.add_argument(
        "--keep-artifacts",
        action="store_true",
        help="leave per-file work dirs in place for inspection",
    )
    p.add_argument(
        "--reference-manifest",
        type=Path,
        action="append",
        default=[],
        help=(
            "reference manifest JSON (Spec 034) enabling per-feature "
            "survival probes for the matching reference document; "
            "repeatable"
        ),
    )
    args = p.parse_args()

    try:
        find_soffice()
    except SofficeNotFoundError as exc:
        print(f"soffice unavailable: {exc}", file=sys.stderr)
        return 2

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(
                sorted(
                    p
                    for p in entry.rglob("*")
                    if p.is_file() and p.suffix.lower() in _FORMAT_BY_EXT
                )
            )
        elif entry.is_file():
            inputs.append(entry)

    if not inputs:
        print("no OOXML inputs found", file=sys.stderr)
        return 2

    reference_manifests: dict[str, dict] = {}
    for manifest_path in args.reference_manifest:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        document_name = manifest_path.name.removesuffix(".manifest.json")
        reference_manifests[document_name] = payload

    args.work_root.mkdir(parents=True, exist_ok=True)
    observations = observe_batch(
        inputs,
        args.work_root,
        timeout_seconds=args.timeout,
        keep_artifacts=args.keep_artifacts,
        reference_manifests=reference_manifests,
    )
    report = _to_jsonable(observations)

    if args.output:
        args.output.write_text(json.dumps(report, indent=2))
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(json.dumps(report, indent=2))

    summary = report["summary"]
    print(
        f"\nlo-ooxml-oracle: total={summary['total']} "
        f"preserved={summary['preserved']} repaired={summary['repaired']} "
        f"crash={summary['crash']} timeout={summary['timeout']} "
        f"open_failed={summary['open_failed']}",
        file=sys.stderr,
    )
    hard_errors = (
        summary["crash"] + summary["timeout"] + summary["open_failed"] + summary["missing_output"]
    )
    return 1 if hard_errors > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
