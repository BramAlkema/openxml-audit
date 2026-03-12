#!/usr/bin/env python3
"""Run an ODF parity snapshot against the pinned corpus manifest.

Mirrors the OOXML parity snapshot pattern: each corpus sample has a pinned
profile (valid/invalid) and an optional expected_error_count.  The Python
ODF validator runs on materialized fixtures and results are compared against
those expectations.  No external Java tools required.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import FileFormat, OdfValidator  # noqa: E402
from openxml_audit.parity_normalization import (  # noqa: E402
    normalize_error_tuple,
)

DEFAULT_CORPUS_MANIFEST = Path("data/odf/reference_corpus/manifest.json")
DEFAULT_FIXTURES_ROOT = Path("tests/fixtures/odf")
DEFAULT_OUTPUT = Path("data/odf/parity_baseline/parity_snapshot.json")

FILE_FORMAT_MAP: dict[str, FileFormat] = {
    "odf1.2": FileFormat.ODF_1_2,
    "odf1.3": FileFormat.ODF_1_3,
}


def _materialize_odf(fixture_dir: Path, output_path: Path) -> None:
    """Build a deterministic ODF ZIP from a fixture directory."""
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        mimetype_path = fixture_dir / "mimetype"
        if mimetype_path.exists():
            zf.write(mimetype_path, "mimetype", compress_type=zipfile.ZIP_STORED)

        for child in sorted(fixture_dir.rglob("*")):
            if child.is_file() and child.name != "mimetype":
                arcname = str(child.relative_to(fixture_dir))
                zf.write(child, arcname)


def main() -> int:
    wall_start = perf_counter()
    parser = argparse.ArgumentParser(
        description="Run ODF parity snapshot against corpus expectations."
    )
    parser.add_argument(
        "--corpus-manifest",
        type=Path,
        default=DEFAULT_CORPUS_MANIFEST,
        help=f"Corpus manifest path (default: {DEFAULT_CORPUS_MANIFEST})",
    )
    parser.add_argument(
        "--fixtures-root",
        type=Path,
        default=DEFAULT_FIXTURES_ROOT,
        help=f"Fixture root (default: {DEFAULT_FIXTURES_ROOT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Snapshot output path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--strict",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Use strict mode in Python validator.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print summary only without writing report.",
    )
    args = parser.parse_args()

    manifest_path = args.corpus_manifest.resolve()
    if not manifest_path.exists():
        print(f"Manifest not found: {manifest_path}")
        return 2
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    samples = manifest.get("samples", [])
    if not samples:
        print("No samples in manifest.")
        return 2

    fixtures_root = args.fixtures_root.resolve()
    by_category: Counter[str] = Counter()
    by_profile: Counter[str] = Counter()
    mismatch_families: Counter[str] = Counter()
    family_details: dict[str, dict[str, str]] = {}
    mismatch_examples: list[dict[str, Any]] = []
    matched = 0
    mismatched = 0

    with TemporaryDirectory(prefix="odf_parity_") as tmpdir:
        staging = Path(tmpdir)

        for sample in samples:
            sample_id = sample["id"]
            profile = sample["profile"]
            category = sample.get("category", "unknown")
            fixture_dir = fixtures_root / sample["fixture_dir"]
            filename = sample["filename"]
            file_format_str = sample.get("file_format", "odf1.3")
            file_format = FILE_FORMAT_MAP.get(file_format_str, FileFormat.ODF_1_3)

            by_category[category] += 1
            by_profile[profile] += 1

            odf_path = staging / sample_id / filename
            odf_path.parent.mkdir(parents=True, exist_ok=True)
            _materialize_odf(fixture_dir, odf_path)

            validator = OdfValidator(
                file_format=file_format, strict=args.strict, security_validation=True
            )
            result = validator.validate(odf_path)
            error_count = len(result.errors)

            # Determine expected error count: explicit or from profile.
            expected = sample.get("expected_error_count")
            if expected is not None:
                ok = error_count == expected
            elif profile == "valid":
                ok = error_count == 0
            else:
                ok = error_count > 0

            if ok:
                matched += 1
            else:
                mismatched += 1
                mismatch_examples.append(
                    {
                        "sample_id": sample_id,
                        "profile": profile,
                        "category": category,
                        "expected": expected if expected is not None else (
                            "0 (valid)" if profile == "valid" else ">0 (invalid)"
                        ),
                        "actual_error_count": error_count,
                        "errors": [e.description for e in result.errors[:10]],
                    }
                )
                for error in result.errors[:20]:
                    normalized = normalize_error_tuple(error)
                    fk = normalized.get("family_key", "")
                    if fk:
                        mismatch_families[fk] += 1
                        if fk not in family_details:
                            family_details[fk] = normalized

    total = matched + mismatched
    match_rate = 100.0 * matched / total if total else 0.0
    duration = perf_counter() - wall_start

    report = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "manifest": str(manifest_path),
        "fixtures_root": str(fixtures_root),
        "strict": args.strict,
        "duration_seconds": round(duration, 6),
        "checks_total": total,
        "checks_matched": matched,
        "checks_mismatched": mismatched,
        "match_rate_percent": round(match_rate, 2),
        "by_profile": dict(sorted(by_profile.items())),
        "by_category": dict(sorted(by_category.items())),
        "mismatch_examples": mismatch_examples,
        "mismatch_families": [
            {
                **family_details.get(
                    key,
                    {"description": key, "family_key": key},
                ),
                "count": count,
            }
            for key, count in mismatch_families.most_common(50)
        ],
    }

    print(f"Checks total: {total}")
    print(f"Matched: {matched}")
    print(f"Mismatched: {mismatched}")
    print(f"Match rate: {report['match_rate_percent']}%")
    print(f"Duration: {report['duration_seconds']}s")
    if report["mismatch_families"]:
        top = report["mismatch_families"][0]
        print(f"Top mismatch family: {top['count']}x {top.get('description', '')}")

    if args.dry_run:
        print("Dry run — report not written.")
        return 0

    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Wrote report: {output}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
