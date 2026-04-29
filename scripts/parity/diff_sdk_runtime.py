#!/usr/bin/env python3
"""Diff our Python validator's output against the .NET SDK runtime output.

The .NET SDK runtime snapshot is produced by
`tools/parity/dotnet_validator_runner` (a small .NET CLI that walks the
corpus, invokes `OpenXmlValidator` at each requested FileFormatVersion,
and emits per-file JSON results). This script runs our Python validator
on the same corpus and diffs the family-key sets, per (file, version).

Spec 013 OQ8 option B ("live SDK runtime as the parity anchor") is the
intended consumer; this script is the proof-of-concept comparator.

Output: a markdown table per file with deltas. Files with no delta are
reported as "OK" rows; non-empty deltas list the family-keys that
appear in one snapshot but not the other.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

# Strip XML namespace prefixes from path segments (e.g. /w:document[1] -> /document[1]).
# The .NET SDK emits paths with the source-document's namespace prefixes preserved
# (w:, mc:, v:, ...); our Python validator emits paths with prefixes stripped.
# To compare apples to apples, normalize both sides through the same regex.
_NS_PREFIX_RE = re.compile(r"/[a-zA-Z][a-zA-Z0-9]*:")


def strip_ns_prefixes(path: str) -> str:
    return _NS_PREFIX_RE.sub("/", path) if path else path

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC = REPO_ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from openxml_audit.errors import ValidationError, ValidationErrorType
from openxml_audit.parity_normalization import normalize_error_tuple
from openxml_audit.validator import OpenXmlValidator
from openxml_audit.errors import FileFormat


PY_FORMAT_BY_NAME = {
    "Office2007": FileFormat.OFFICE_2007,
    "Office2010": FileFormat.OFFICE_2010,
    "Office2013": FileFormat.OFFICE_2013,
    "Office2016": FileFormat.OFFICE_2016,
    "Microsoft365": FileFormat.MICROSOFT_365,
}


@dataclass
class FamilyDelta:
    file: str
    version: str
    py_count: int
    sdk_count: int
    only_py: list[str]
    only_sdk: list[str]
    common_count_diffs: dict[str, tuple[int, int]]


def load_sdk_runtime(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def family_key_from_sdk_record(rec: dict) -> str:
    err = ValidationError(
        id=rec["id"],
        error_type=_map_sdk_error_type(rec["error_type"]),
        description=rec["description"],
        path=strip_ns_prefixes(rec["path"]),
        part_uri=rec["part"],
    )
    return normalize_error_tuple(err)["family_key"]


def _map_sdk_error_type(name: str) -> ValidationErrorType:
    mapping = {
        "Schema": ValidationErrorType.SCHEMA,
        "Semantic": ValidationErrorType.SEMANTIC,
        "Package": ValidationErrorType.PACKAGE,
        "MarkupCompatibility": ValidationErrorType.MARKUP_COMPATIBILITY,
    }
    return mapping.get(name, ValidationErrorType.SCHEMA)


def py_family_keys(file_path: Path, version_name: str) -> list[str]:
    fmt = PY_FORMAT_BY_NAME[version_name]
    validator = OpenXmlValidator(file_format=fmt, strict=True)
    result = validator.validate(file_path)
    return [normalize_error_tuple(e)["family_key"] for e in result.errors]


def diff_one(file_path: Path, relpath: str, sdk_versions: list[dict]) -> list[FamilyDelta]:
    deltas: list[FamilyDelta] = []
    for sdk_v in sdk_versions:
        version = sdk_v["version"]
        if version not in PY_FORMAT_BY_NAME:
            continue
        if sdk_v.get("open_error"):
            continue
        sdk_keys = Counter(family_key_from_sdk_record(e) for e in sdk_v["errors"])
        py_keys = Counter(py_family_keys(file_path, version))
        only_py = sorted(set(py_keys) - set(sdk_keys))
        only_sdk = sorted(set(sdk_keys) - set(py_keys))
        common_diffs = {
            k: (py_keys[k], sdk_keys[k])
            for k in (set(py_keys) & set(sdk_keys))
            if py_keys[k] != sdk_keys[k]
        }
        deltas.append(FamilyDelta(
            file=relpath,
            version=version,
            py_count=sum(py_keys.values()),
            sdk_count=sum(sdk_keys.values()),
            only_py=only_py,
            only_sdk=only_sdk,
            common_count_diffs=common_diffs,
        ))
    return deltas


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--sdk-runtime", required=True, type=Path,
                   help="path to SDK runtime snapshot JSON")
    p.add_argument("--files-root", required=True, type=Path,
                   help="corpus root (same as the one passed to the .NET runner)")
    p.add_argument("--filter", default=None,
                   help="optional substring filter on source_relpath")
    p.add_argument("--only-deltas", action="store_true",
                   help="only print files that have at least one delta")
    args = p.parse_args()

    snap = load_sdk_runtime(args.sdk_runtime)
    files = snap["files"]
    if args.filter:
        files = [f for f in files if args.filter in f["source_relpath"]]

    has_any_delta = False
    for f in files:
        relpath = f["source_relpath"]
        file_path = args.files_root / relpath
        if not file_path.exists():
            print(f"SKIP {relpath}: not found at {file_path}")
            continue
        deltas = diff_one(file_path, relpath, f["validations"])
        any_delta = any(
            d.only_py or d.only_sdk or d.common_count_diffs
            for d in deltas
        )
        if args.only_deltas and not any_delta:
            continue
        if any_delta:
            has_any_delta = True
        print(f"\n## {relpath}")
        for d in deltas:
            label = "OK" if not (d.only_py or d.only_sdk or d.common_count_diffs) else "DELTA"
            print(f"  [{label}] {d.version:14s} py={d.py_count:5d} sdk={d.sdk_count:5d}")
            if d.only_py:
                print(f"      only-py ({len(d.only_py)}):")
                for k in d.only_py[:10]:
                    print(f"        + {k}")
                if len(d.only_py) > 10:
                    print(f"        ... and {len(d.only_py) - 10} more")
            if d.only_sdk:
                print(f"      only-sdk ({len(d.only_sdk)}):")
                for k in d.only_sdk[:10]:
                    print(f"        - {k}")
                if len(d.only_sdk) > 10:
                    print(f"        ... and {len(d.only_sdk) - 10} more")
            if d.common_count_diffs:
                print(f"      common-with-count-diff ({len(d.common_count_diffs)}):")
                for k, (py, sdk) in list(d.common_count_diffs.items())[:10]:
                    print(f"        ~ py={py} sdk={sdk}  {k}")

    return 1 if has_any_delta else 0


if __name__ == "__main__":
    raise SystemExit(main())
