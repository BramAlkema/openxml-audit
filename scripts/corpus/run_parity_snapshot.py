#!/usr/bin/env python3
"""Run a parity snapshot against manifest expectations."""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from time import perf_counter
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import FileFormat, OpenXmlValidator  # noqa: E402
from openxml_audit.parity_normalization import (  # noqa: E402
    normalize_description,
    normalize_error_tuple,
)

DEFAULT_MANIFEST = Path("data/corpus/sdk_seed/manifest.json")
DEFAULT_OUTPUT = Path("data/corpus/parity_baseline/v3.4.1/parity_snapshot.json")

VERSION_MAP = {
    "Office2007": FileFormat.OFFICE_2007,
    "Office2010": FileFormat.OFFICE_2010,
    "Office2013": FileFormat.OFFICE_2013,
    "Office2016": FileFormat.OFFICE_2016,
    "Office2019": FileFormat.OFFICE_2019,
}


@dataclass
class ValidationRun:
    """Cached validation run for one file/version."""

    error_count: int
    descriptions: list[str]
    normalized_tuples: list[dict[str, str]]


def _load_manifest(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _expected_from_entry(entry: dict[str, Any]) -> tuple[int | None, list[int]]:
    if "expected_error_count" in entry:
        value = entry["expected_error_count"]
        if isinstance(value, int):
            return value, []
    counts = entry.get("expected_error_counts")
    if isinstance(counts, list):
        values = sorted({int(v) for v in counts if isinstance(v, int)})
        return None, values
    return None, []


def _get_versions(entry: dict[str, Any]) -> list[str]:
    versions = entry.get("validator_versions")
    if isinstance(versions, list) and versions:
        return [str(v) for v in versions if str(v) in VERSION_MAP]
    single = entry.get("file_format_version")
    if isinstance(single, str) and single in VERSION_MAP:
        return [single]
    return ["Office2007"]


def _run_validator(
    cache: dict[tuple[str, str], ValidationRun],
    file_path: Path,
    version: str,
    strict: bool,
) -> ValidationRun:
    key = (str(file_path), version)
    if key in cache:
        return cache[key]
    validator = OpenXmlValidator(
        file_format=VERSION_MAP[version],
        max_errors=0,
        security_validation=False,
        strict=strict,
    )
    result = validator.validate(file_path)
    run = ValidationRun(
        error_count=len(result.errors),
        descriptions=[error.description for error in result.errors],
        normalized_tuples=[normalize_error_tuple(error) for error in result.errors],
    )
    cache[key] = run
    return run


def _collect_checks(
    manifest: dict[str, Any],
    max_files: int,
    max_checks: int,
    include_mutation_expectations: bool,
) -> tuple[list[dict[str, Any]], int]:
    checks: list[dict[str, Any]] = []
    skipped_mutation = 0
    files = manifest.get("files", [])
    if not isinstance(files, list):
        return checks, skipped_mutation

    processed_files = 0
    for file_entry in files:
        expectations = file_entry.get("expectations")
        relpath = file_entry.get("source_relpath")
        if not isinstance(relpath, str) or not isinstance(expectations, list):
            continue
        if not expectations:
            continue
        processed_files += 1
        for expectation in expectations:
            if not isinstance(expectation, dict):
                continue
            if (
                str(expectation.get("scenario", "base")) == "mutation"
                and not include_mutation_expectations
            ):
                skipped_mutation += 1
                continue
            checks.append(
                {
                    "source_relpath": relpath,
                    "expectation": expectation,
                }
            )
            if max_checks > 0 and len(checks) >= max_checks:
                return checks, skipped_mutation
        if max_files > 0 and processed_files >= max_files:
            break
    return checks, skipped_mutation


def main() -> int:
    wall_start = perf_counter()
    parser = argparse.ArgumentParser(description="Run parity snapshot from manifest expectations.")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=DEFAULT_MANIFEST,
        help=f"Manifest path (default: {DEFAULT_MANIFEST})",
    )
    parser.add_argument(
        "--files-root",
        type=Path,
        default=None,
        help="Root for source_relpath files (default: manifest output_root + '/files')",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Snapshot output path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Maximum files with expectations to include (0 means no limit).",
    )
    parser.add_argument(
        "--max-checks",
        type=int,
        default=0,
        help="Maximum expectation checks to run (0 means no limit).",
    )
    parser.add_argument(
        "--max-mismatches",
        type=int,
        default=200,
        help="Maximum mismatch examples stored in report.",
    )
    parser.add_argument(
        "--include-mutation-expectations",
        action="store_true",
        help="Include mutation-dependent expectations extracted from SDK tests.",
    )
    parser.add_argument(
        "--strict",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Use strict mode in Python validator.",
    )
    parser.add_argument(
        "--allow-missing-files",
        action="store_true",
        help="Do not fail when corpus files referenced by manifest are missing.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print summary only without writing report.",
    )
    args = parser.parse_args()

    manifest_path = args.manifest.resolve()
    if not manifest_path.exists():
        print(f"Manifest not found: {manifest_path}")
        return 2
    manifest = _load_manifest(manifest_path)

    if args.files_root is not None:
        files_root = args.files_root.resolve()
    else:
        output_root = manifest.get("output_root")
        if not isinstance(output_root, str):
            print("Manifest missing 'output_root'; provide --files-root.")
            return 2
        files_root = (Path(output_root) / "files").resolve()

    checks, skipped_mutation_expectations = _collect_checks(
        manifest,
        max_files=args.max_files,
        max_checks=args.max_checks,
        include_mutation_expectations=args.include_mutation_expectations,
    )
    if not checks:
        print("No expectation checks found.")
        return 2

    cache: dict[tuple[str, str], ValidationRun] = {}
    by_kind = Counter()
    by_version = Counter()
    mismatch_families = Counter()
    family_details: dict[str, dict[str, str]] = {}
    mismatch_examples: list[dict[str, Any]] = []
    evaluated_checks = 0
    mismatched_checks = 0
    matched = 0
    missing_file_checks = 0

    for check in checks:
        relpath = check["source_relpath"]
        expectation = check["expectation"]
        kind = str(expectation.get("kind", "unknown"))
        by_kind[kind] += 1

        versions = _get_versions(expectation)
        for version in versions:
            by_version[version] += 1

        file_path = files_root / relpath
        if not file_path.exists():
            missing_file_checks += 1
            evaluated_checks += 1
            mismatched_checks += 1
            mismatch_families["file_not_found"] += 1
            if len(mismatch_examples) < args.max_mismatches:
                mismatch_examples.append(
                    {
                        "source_relpath": relpath,
                        "kind": kind,
                        "validator_versions": versions,
                        "error": "file_not_found",
                    }
                )
            continue

        actual_by_version: dict[str, int] = {}
        descriptions_by_version: dict[str, list[str]] = {}
        tuples_by_version: dict[str, list[dict[str, str]]] = {}
        for version in versions:
            run = _run_validator(
                cache=cache,
                file_path=file_path,
                version=version,
                strict=args.strict,
            )
            actual_by_version[version] = run.error_count
            descriptions_by_version[version] = run.descriptions
            tuples_by_version[version] = run.normalized_tuples

        actual = sum(actual_by_version.values())
        expected_count, expected_counts = _expected_from_entry(expectation)
        if expected_count is not None:
            ok = actual == expected_count
            expected_payload: Any = expected_count
        elif expected_counts:
            ok = actual in expected_counts
            expected_payload = expected_counts
        else:
            # No usable expectation payload; skip from denominator.
            continue

        evaluated_checks += 1
        if ok:
            matched += 1
            continue

        mismatched_checks += 1
        if len(mismatch_examples) < args.max_mismatches:
            mismatch_examples.append(
                {
                    "source_relpath": relpath,
                    "kind": kind,
                    "validator_versions": versions,
                    "expected": expected_payload,
                    "actual_sum": actual,
                    "actual_by_version": actual_by_version,
                }
            )

        for version in versions:
            top_tuples = tuples_by_version.get(version, [])[:20]
            for normalized_tuple in top_tuples:
                family_key = normalized_tuple.get("family_key", "")
                if family_key:
                    mismatch_families[family_key] += 1
                    if family_key not in family_details:
                        family_details[family_key] = normalized_tuple
                    continue

                # Fallback for unexpected payload shape.
                description = normalized_tuple.get("description", "")
                if not description:
                    continue
                fallback_key = f"unknown|unknown|/|/|{normalize_description(description)}"
                mismatch_families[fallback_key] += 1
                if fallback_key not in family_details:
                    family_details[fallback_key] = {
                        "id": "unknown",
                        "error_type": "unknown",
                        "part": "/",
                        "path": "/",
                        "description": normalize_description(description),
                        "family_key": fallback_key,
                    }

    total = evaluated_checks
    mismatch_count = mismatched_checks
    match_rate = 100.0 * matched / total if total else 0.0
    duration_seconds = perf_counter() - wall_start

    report = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "manifest": str(manifest_path),
        "files_root": str(files_root),
        "strict": args.strict,
        "allow_missing_files": args.allow_missing_files,
        "include_mutation_expectations": args.include_mutation_expectations,
        "skipped_mutation_expectations": skipped_mutation_expectations,
        "duration_seconds": round(duration_seconds, 6),
        "checks_collected": len(checks),
        "checks_total": total,
        "checks_matched": matched,
        "checks_mismatched": mismatch_count,
        "checks_missing_files": missing_file_checks,
        "match_rate_percent": round(match_rate, 2),
        "validation_runs": len(cache),
        "by_kind": dict(sorted(by_kind.items())),
        "by_version": dict(sorted(by_version.items())),
        "mismatch_examples": mismatch_examples,
        "mismatch_families": [
            {
                **family_details.get(
                    key,
                    {
                        "id": "unknown",
                        "error_type": "unknown",
                        "part": "/",
                        "path": "/",
                        "description": key,
                        "family_key": key,
                    },
                ),
                "count": count,
            }
            for key, count in mismatch_families.most_common(50)
        ],
    }

    print(f"Checks total: {total}")
    print(f"Matched: {matched}")
    print(f"Mismatched: {mismatch_count}")
    if missing_file_checks:
        print(f"Missing files: {missing_file_checks}")
    print(f"Match rate: {report['match_rate_percent']}%")
    print(f"Duration: {report['duration_seconds']}s")
    print(f"Validation runs: {report['validation_runs']}")
    if report["mismatch_families"]:
        top = report["mismatch_families"][0]
        print(f"Top mismatch family: {top['count']}x {top['description']}")

    if missing_file_checks and not args.allow_missing_files:
        print(
            "Missing corpus files detected. Materialize corpus files first or pass "
            "--allow-missing-files."
        )
        if args.dry_run:
            return 2

    if args.dry_run:
        print("Dry run enabled, report not written.")
        return 0

    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Wrote report: {output}")

    if missing_file_checks and not args.allow_missing_files:
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
