"""Word roundtrip oracle, corpus walker.

The existing `tools/oracle/word_repair_oracle.py` is a scenario-matrix
runner (Spec 010 / 011 phase work — generates synthetic .docx files
from a property-element ordering matrix). This driver does the
complementary thing: it walks an arbitrary corpus, calls
`word_roundtrip.roundtrip()` on each file, and emits the same
`RoundtripObservation` shape as the ODF / PowerPoint / Excel oracle
CLIs.

Use it to collect first-use baselines on TokenMoulds-emitted (or
otherwise-real) .docx corpora. Output shape matches the other
oracles — same `_to_jsonable` summary structure, same outcome
vocabulary (`preserved` / `repaired` / `crash` / `timeout` /
`open_failed` / `missing_output`).

Spec: `specs/022-first-oracle-baselines.md`.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import time
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Literal

from openxml_audit.docx import osa as docx_osa  # for app-bundle path / preflight
from openxml_audit.package_diff import compare_packages

# word_roundtrip imports `from tools.oracle import word_window`, which
# requires the repo root (not the tools/ dir) on sys.path. Insert it
# so the import resolves whether or not the caller cd'd to the root.
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from tools.oracle.word_roundtrip import RoundtripError, roundtrip  # noqa: E402


# Canonical OOXML parts to fingerprint. Same shape as
# xlsx_repair_oracle / odf_repair_oracle: skip media, thumbnails, and
# computed-on-save rebuilds.
_FINGERPRINTED_PREFIXES = (
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/stylesWithEffects.xml",
    "word/settings.xml",
    "word/numbering.xml",
    "word/header",
    "word/footer",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/theme/",
    "word/comments.xml",
)


def _is_fingerprinted_part(name: str) -> bool:
    return any(
        name == prefix or (prefix.endswith("/") and name.startswith(prefix))
        or (not prefix.endswith("/") and prefix.endswith(".xml") and name == prefix)
        or (not prefix.endswith("/") and not prefix.endswith(".xml") and name.startswith(prefix))
        for prefix in _FINGERPRINTED_PREFIXES
    )


@dataclass
class PartFingerprint:
    name: str
    sha256: str
    size_bytes: int


@dataclass
class RoundtripObservation:
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
    diff_dir: str | None = None
    repair_dialog_seen: bool = False
    repair_dialog_text: str | None = None
    word_version: str | None = None
    notes: list[str] = field(default_factory=list)


def _fingerprint_docx(path: Path) -> list[PartFingerprint]:
    fingerprints: list[PartFingerprint] = []
    if not path.exists():
        return fingerprints
    with zipfile.ZipFile(path) as zf:
        for name in sorted(zf.namelist()):
            if not _is_fingerprinted_part(name):
                continue
            data = zf.read(name)
            fingerprints.append(PartFingerprint(
                name=name,
                sha256=hashlib.sha256(data).hexdigest(),
                size_bytes=len(data),
            ))
    return fingerprints


def observe(
    input_docx: Path,
    *,
    timeout_seconds: float = 60.0,
    keep_artifacts: bool = False,
) -> RoundtripObservation:
    input_docx = input_docx.resolve()
    if not input_docx.exists():
        return RoundtripObservation(
            source_relpath=str(input_docx.name),
            outcome="missing_output",
            duration_seconds=0.0,
            notes=[f"input does not exist: {input_docx}"],
        )

    input_parts = _fingerprint_docx(input_docx)
    started = time.monotonic()

    try:
        result = roundtrip(input_docx, timeout=timeout_seconds, accept_repair=True)
    except RoundtripError as exc:
        msg = str(exc).lower()
        if "did not register" in msg or "did not open" in msg:
            outcome: Literal[
                "preserved", "repaired", "crash", "timeout",
                "open_failed", "missing_output",
            ] = "open_failed"
        elif "timed out" in msg or "timeout" in msg:
            outcome = "timeout"
        else:
            outcome = "crash"
        return RoundtripObservation(
            source_relpath=str(input_docx.name),
            outcome=outcome,
            duration_seconds=time.monotonic() - started,
            input_parts=input_parts,
            notes=[f"RoundtripError: {exc}"],
        )
    except (RuntimeError, subprocess.TimeoutExpired) as exc:
        return RoundtripObservation(
            source_relpath=str(input_docx.name),
            outcome="crash",
            duration_seconds=time.monotonic() - started,
            input_parts=input_parts,
            notes=[f"{type(exc).__name__}: {exc}"],
        )

    output_parts = _fingerprint_docx(result.output_path)

    # Per-part canonical-c14n diff via the shared package_diff module.
    # Output dir lives next to the post-Word file in the staging tree
    # so callers can `--keep-artifacts` to inspect what Word changed.
    compare_dir = result.output_path.parent / "compare"
    diff_report = compare_packages(
        base_path=input_docx,
        head_path=result.output_path,
        output_dir=compare_dir,
        parts_filter=_is_fingerprinted_part,
    )
    changed = list(diff_report.get("changed_files", []))
    added = list(diff_report.get("added_files", []))
    removed = list(diff_report.get("removed_files", []))

    diff_dir_str: str | None = None
    if keep_artifacts:
        diff_dir_str = str(compare_dir)
    else:
        # word_roundtrip leaves the staging tree in place by default
        # for debuggability; baseline collection doesn't need it.
        if result.output_path.parent != input_docx.parent:
            shutil.rmtree(result.output_path.parent, ignore_errors=True)

    return RoundtripObservation(
        source_relpath=str(input_docx.name),
        outcome="preserved" if not (changed or added or removed) else "repaired",
        duration_seconds=result.elapsed_seconds,
        input_parts=input_parts,
        output_parts=output_parts,
        changed_parts=changed,
        added_parts=added,
        removed_parts=removed,
        diff_dir=diff_dir_str,
        repair_dialog_seen=result.repair_dialog_seen,
        repair_dialog_text=result.repair_dialog_text,
        word_version=result.word_version,
    )


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
                   help=".docx files to roundtrip (or directories to walk)")
    p.add_argument("--output", type=Path, default=None,
                   help="path to write the JSON observation report")
    p.add_argument("--timeout", type=float, default=60.0,
                   help="per-file Word timeout in seconds (default 60)")
    p.add_argument("--keep-artifacts", action="store_true",
                   help="leave staging dirs in place for inspection")
    args = p.parse_args()

    if not docx_osa.WORD_APP_BUNDLE.exists():
        print(
            f"Microsoft Word not installed at {docx_osa.WORD_APP_BUNDLE}",
            file=sys.stderr,
        )
        return 2

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(sorted(entry.rglob("*.docx")))
        elif entry.is_file() and entry.suffix.lower() == ".docx":
            inputs.append(entry)

    if not inputs:
        print("no .docx inputs found", file=sys.stderr)
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
        f"\nword-oracle: total={summary['total']} "
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
