#!/usr/bin/env python3
"""Run a self-parity snapshot of this validator's output.

Self-parity (Spec 013, Spec 026) keys the regression check on **our
own** validator output, not on the .NET SDK's. The snapshot is a
deterministic inventory of every `family_key` (the 5-tuple normalized
by `parity_normalization.normalize_error_tuple`) emitted across the
corpus, along with per-family counts and the `source_class` tag from
Spec 018.

Spec 026's 0.7.1 release ships this as the prototype + an advisory
CI workflow; 0.8.0 promotes the comparator (`compare_self_parity.py`)
to the blocking sovereign gate.

Usage:

    python scripts/parity/run_self_parity_snapshot.py \\
        --manifest data/corpus/sdk_seed/manifest.json \\
        --files-root /tmp/parity-corpus/files \\
        --output data/corpus/self_parity_baseline/v0.7.1/snapshot.json
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import FileFormat, OpenXmlValidator  # noqa: E402
from openxml_audit.errors import SourceClass  # noqa: E402
from openxml_audit.parity_normalization import (  # noqa: E402
    normalize_error_tuple,
)


VERSION_MAP = {
    "Office2007": FileFormat.OFFICE_2007,
    "Office2010": FileFormat.OFFICE_2010,
    "Office2013": FileFormat.OFFICE_2013,
    "Office2016": FileFormat.OFFICE_2016,
    "Office2019": FileFormat.OFFICE_2019,
}


@dataclass
class FamilyEntry:
    """One row in the inventory: a single normalized family_key + how
    often it appeared and what source_class produced it.

    `count` is total occurrences across the whole corpus. The first
    sample fields preserve diagnostic context — the actual unfilled
    description from one of the matching findings, the part it was
    on, and the path the validator emitted. Templated `<value>`
    in the family_key obscures the real attribute name; the sample
    keeps it discoverable.
    """

    family_key: str
    count: int = 0
    source_class: str = SourceClass.SDK_PROXY.value
    first_part: str = ""
    first_path_sample: str = ""
    first_description_sample: str = ""


@dataclass
class SelfParitySnapshot:
    schema_version: int = 1
    generated_at_utc: str = ""
    validator_version: str = ""
    files_root: str = ""
    manifest: str = ""
    file_count: int = 0
    validation_runs: int = 0
    total_findings: int = 0
    by_source_class: dict[str, int] = field(default_factory=dict)
    family_inventory: dict[str, FamilyEntry] = field(default_factory=dict)


def _load_manifest(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _file_format_versions(entry: dict[str, Any]) -> list[str]:
    versions = entry.get("file_format_versions") or []
    if isinstance(versions, list) and versions:
        return [str(v) for v in versions if str(v) in VERSION_MAP]
    single = entry.get("file_format_version")
    if isinstance(single, str) and single in VERSION_MAP:
        return [single]
    return ["Office2007"]


def _gather_findings(
    files: list[dict[str, Any]],
    files_root: Path,
    *,
    strict: bool = True,
) -> tuple[dict[str, FamilyEntry], dict[str, int], int, int]:
    """Walk the corpus, validate each file at each FileFormat, and
    accumulate the inventory.

    Returns (family_inventory, source_class_counts, validation_runs,
    total_findings).
    """
    inventory: dict[str, FamilyEntry] = {}
    source_class_counts: dict[str, int] = defaultdict(int)
    runs = 0
    total = 0

    for entry in files:
        rel = entry.get("source_relpath")
        if not rel:
            continue
        file_path = files_root / rel
        if not file_path.exists():
            continue
        for version in _file_format_versions(entry):
            validator = OpenXmlValidator(
                file_format=VERSION_MAP[version],
                max_errors=0,
                security_validation=False,
                strict=strict,
            )
            result = validator.validate(file_path)
            runs += 1
            for err in result.errors:
                normalized = normalize_error_tuple(err)
                key = normalized["family_key"]
                source = err.source_class.value if err.source_class else SourceClass.SDK_PROXY.value
                source_class_counts[source] += 1
                total += 1
                slot = inventory.get(key)
                if slot is None:
                    inventory[key] = FamilyEntry(
                        family_key=key,
                        count=1,
                        source_class=source,
                        first_part=err.part_uri,
                        first_path_sample=err.path,
                        first_description_sample=err.description,
                    )
                else:
                    slot.count += 1

    return inventory, dict(source_class_counts), runs, total


def write_snapshot(
    *,
    manifest_path: Path,
    files_root: Path,
    output: Path,
    validator_version: str,
    strict: bool = True,
) -> SelfParitySnapshot:
    manifest = _load_manifest(manifest_path)
    files = manifest.get("files") or []
    inventory, source_classes, runs, total = _gather_findings(
        files, files_root, strict=strict,
    )

    snapshot = SelfParitySnapshot(
        generated_at_utc=datetime.now(timezone.utc).isoformat(),
        validator_version=validator_version,
        files_root=str(files_root),
        manifest=str(manifest_path),
        file_count=sum(1 for f in files if (files_root / f.get("source_relpath", "")).exists()),
        validation_runs=runs,
        total_findings=total,
        by_source_class=dict(sorted(source_classes.items())),
        family_inventory={k: inventory[k] for k in sorted(inventory)},
    )

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(
            {
                "schema_version": snapshot.schema_version,
                "generated_at_utc": snapshot.generated_at_utc,
                "validator_version": snapshot.validator_version,
                "files_root": snapshot.files_root,
                "manifest": snapshot.manifest,
                "file_count": snapshot.file_count,
                "validation_runs": snapshot.validation_runs,
                "total_findings": snapshot.total_findings,
                "by_source_class": snapshot.by_source_class,
                "family_inventory": {k: asdict(v) for k, v in snapshot.family_inventory.items()},
            },
            indent=2,
            sort_keys=False,
        ),
        encoding="utf-8",
    )
    return snapshot


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--manifest", required=True, type=Path,
                   help="path to corpus manifest (data/corpus/sdk_seed/manifest.json)")
    p.add_argument("--files-root", required=True, type=Path,
                   help="root directory containing the corpus files")
    p.add_argument("--output", required=True, type=Path,
                   help="path to write the self-parity snapshot JSON")
    p.add_argument("--validator-version", default="",
                   help="optional validator version label "
                        "(e.g. '0.7.1'); recorded in the snapshot for traceability")
    p.add_argument("--lax", action="store_true",
                   help="run validator in non-strict mode (errors → warnings)")
    args = p.parse_args()

    if not args.manifest.exists():
        print(f"manifest not found: {args.manifest}", file=sys.stderr)
        return 2
    if not args.files_root.exists():
        print(f"files-root not found: {args.files_root}", file=sys.stderr)
        return 2

    validator_version = args.validator_version
    if not validator_version:
        try:
            from openxml_audit import __version__
            validator_version = __version__
        except Exception:
            validator_version = "unknown"

    snapshot = write_snapshot(
        manifest_path=args.manifest,
        files_root=args.files_root,
        output=args.output,
        validator_version=validator_version,
        strict=not args.lax,
    )

    print(
        f"self-parity snapshot: {snapshot.file_count} files, "
        f"{snapshot.validation_runs} validation runs, "
        f"{snapshot.total_findings} findings, "
        f"{len(snapshot.family_inventory)} unique family keys",
        file=sys.stderr,
    )
    print(f"by source_class: {snapshot.by_source_class}", file=sys.stderr)
    print(f"wrote {args.output}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
