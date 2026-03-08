#!/usr/bin/env python3
"""Extract expected validation outcomes from Open XML SDK tests.

This script parses SDK test sources and updates a corpus manifest with
expectation entries keyed by `source_relpath`.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_SDK_ROOT = Path("/tmp/openxml-sdk-upstream")
DEFAULT_MANIFEST = Path("data/corpus/sdk_seed/manifest.json")
ASSETS_DIR = Path("test/DocumentFormat.OpenXml.Tests.Assets")
TESTS_DIR = Path("test/DocumentFormat.OpenXml.Tests")

CLASS_RE = re.compile(
    r"^\s*public\s+(?:static\s+)?(?:partial\s+)?class\s+([A-Za-z0-9_]+)\s*(?:[:{]|$)"
)
CONST_RE = re.compile(r'^\s*public const string ([A-Za-z0-9_]+)\s*=\s*"([^"]+)";')
METHOD_RE = re.compile(
    r"^\s*public\s+(?:static\s+)?(?:async\s+)?(?:void|Task(?:<[^>]+>)?)\s+([A-Za-z0-9_]+)\s*\("
)
GET_STREAM_RE = re.compile(r"GetStream\(\s*([A-Za-z0-9_\.]+)\s*(?:,|\))")
VALIDATOR_VERSION_RE = re.compile(
    r"new\s+OpenXmlValidator\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*\)"
)
VALIDATOR_DEFAULT_RE = re.compile(r"new\s+OpenXmlValidator\(\s*\)")
INLINE_FILE_COUNT_RE = re.compile(r"\[InlineData\(\s*([A-Za-z0-9_\.]+)\s*,\s*(-?\d+)\s*\)\]")
INLINE_VERSION_COUNT_RE = re.compile(
    r"\[InlineData\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*,\s*(-?\d+)\s*\)\]"
)
INLINE_FILE_ONLY_RE = re.compile(r"\[InlineData\(\s*([A-Za-z0-9_\.]+)\s*\)\]")
INLINE_VERSION_ONLY_RE = re.compile(r"\[InlineData\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*\)\]")
ASSERT_TRUE_ALLOWED_COUNTS_RE = re.compile(
    r"Assert\.True\(\s*cnt\s*==\s*(-?\d+)\s*\|\|\s*cnt\s*==\s*(-?\d+)\s*\)"
)
XLSX_HELPER_CALL_RE = re.compile(
    r"XlsxValidationHelper\(\s*stream\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)"
)
VALID_VERSIONS = {"Office2007", "Office2010", "Office2013", "Office2016", "Office2019"}


@dataclass(frozen=True)
class MethodBlock:
    """Parsed C# method block."""

    name: str
    start_line: int
    attributes: str
    body: str
    source: str


@dataclass(frozen=True)
class ExtractedExpectation:
    """Expectation mapped to a corpus file."""

    relpath: str
    kind: str
    validator_versions: tuple[str, ...]
    expected_error_count: int | None
    expected_error_counts: tuple[int, ...]
    source_file: str
    source_method: str
    source_line: int


def _load_manifest(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_manifest(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _build_relpath_indexes(manifest: dict[str, Any]) -> tuple[dict[str, str], dict[str, str]]:
    exact: dict[str, str] = {}
    lower: dict[str, str] = {}
    for entry in manifest.get("files", []):
        relpath = entry.get("source_relpath")
        if not isinstance(relpath, str):
            continue
        dotted = relpath.replace("/", ".")
        exact[dotted] = relpath
        lower[dotted.lower()] = relpath
    return exact, lower


def _resolve_resource_to_relpath(
    resource: str, dotted_exact: dict[str, str], dotted_lower: dict[str, str]
) -> str | None:
    if resource in dotted_exact:
        return dotted_exact[resource]
    lowered = resource.lower()
    if lowered in dotted_lower:
        return dotted_lower[lowered]

    parts = resource.split(".")
    if len(parts) < 2:
        return None

    for split_index in range(len(parts) - 1, 0, -1):
        directory = "/".join(parts[:split_index])
        filename = ".".join(parts[split_index:])
        candidate = f"{directory}/{filename}"
        dotted = candidate.replace("/", ".")
        if dotted in dotted_exact:
            return dotted_exact[dotted]
        if dotted.lower() in dotted_lower:
            return dotted_lower[dotted.lower()]
    return None


def _parse_constant_symbols(assets_dir: Path) -> dict[str, str]:
    symbol_to_resource: dict[str, str] = {}
    unique_short: dict[str, str] = {}
    ambiguous_short: set[str] = set()

    for path in sorted(assets_dir.glob("TestAssets*.cs")):
        lines = path.read_text(encoding="utf-8").splitlines()
        scope: list[str] = []
        pending_classes: list[str] = []

        for line in lines:
            class_match = CLASS_RE.match(line)
            if class_match is not None:
                pending_classes.append(class_match.group(1))

            const_match = CONST_RE.match(line)
            if const_match is not None:
                const_name, resource = const_match.group(1), const_match.group(2)
                if "TestAssets" in scope:
                    idx = scope.index("TestAssets")
                    prefix = scope[idx + 1 :]
                else:
                    prefix = scope[:]
                symbol = ".".join([*prefix, const_name]) if prefix else const_name
                symbol_to_resource[symbol] = resource

                if const_name in ambiguous_short:
                    pass
                elif const_name in unique_short and unique_short[const_name] != resource:
                    ambiguous_short.add(const_name)
                    unique_short.pop(const_name, None)
                else:
                    unique_short[const_name] = resource

            opens = line.count("{")
            closes = line.count("}")
            while opens > 0 and pending_classes:
                scope.append(pending_classes.pop(0))
                opens -= 1
            for _ in range(closes):
                if scope:
                    scope.pop()

    for short_name, resource in unique_short.items():
        symbol_to_resource.setdefault(short_name, resource)
    return symbol_to_resource


def _find_method_end(lines: list[str], start_index: int) -> int:
    open_line = -1
    open_col = -1
    for i in range(start_index, len(lines)):
        col = lines[i].find("{")
        if col >= 0:
            open_line = i
            open_col = col
            break
    if open_line < 0:
        return start_index

    depth = 0
    for i in range(open_line, len(lines)):
        start_col = open_col if i == open_line else 0
        line = lines[i]
        for col in range(start_col, len(line)):
            char = line[col]
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    return i
    return len(lines) - 1


def _collect_attributes(lines: list[str], method_start: int) -> str:
    attrs: list[str] = []
    i = method_start - 1
    while i >= 0:
        stripped = lines[i].strip()
        if stripped.startswith("["):
            attrs.append(lines[i])
            i -= 1
            continue
        if not stripped:
            i -= 1
            continue
        break
    attrs.reverse()
    return "\n".join(attrs)


def _iter_methods(test_path: Path) -> list[MethodBlock]:
    lines = test_path.read_text(encoding="utf-8").splitlines()
    methods: list[MethodBlock] = []
    i = 0
    while i < len(lines):
        match = METHOD_RE.match(lines[i])
        if match is None:
            i += 1
            continue
        name = match.group(1)
        end = _find_method_end(lines, i)
        methods.append(
            MethodBlock(
                name=name,
                start_line=i + 1,
                attributes=_collect_attributes(lines, i),
                body="\n".join(lines[i : end + 1]),
                source=str(test_path),
            )
        )
        i = end + 1
    return methods


def _normalize_versions(versions: list[str]) -> tuple[str, ...]:
    ordered: list[str] = []
    seen: set[str] = set()
    for version in versions:
        if version not in VALID_VERSIONS:
            continue
        if version in seen:
            continue
        seen.add(version)
        ordered.append(version)
    return tuple(ordered)


def _resolve_symbol(
    symbol: str,
    symbol_to_resource: dict[str, str],
    dotted_exact: dict[str, str],
    dotted_lower: dict[str, str],
) -> str | None:
    candidates = [symbol]
    if symbol.startswith("TestAssets."):
        candidates.append(symbol.removeprefix("TestAssets."))
    if "." in symbol:
        candidates.append(symbol.split(".")[-1])

    for candidate in candidates:
        resource = symbol_to_resource.get(candidate)
        if resource is None:
            continue
        relpath = _resolve_resource_to_relpath(resource, dotted_exact, dotted_lower)
        if relpath is not None:
            return relpath
    return None


def _build_expectation(
    relpath: str,
    kind: str,
    versions: tuple[str, ...],
    source: MethodBlock,
    expected_count: int | None = None,
    expected_counts: tuple[int, ...] = (),
) -> ExtractedExpectation:
    return ExtractedExpectation(
        relpath=relpath,
        kind=kind,
        validator_versions=versions,
        expected_error_count=expected_count,
        expected_error_counts=expected_counts,
        source_file=Path(source.source).name,
        source_method=source.name,
        source_line=source.start_line,
    )


def _extract_from_method(
    method: MethodBlock,
    symbol_to_resource: dict[str, str],
    dotted_exact: dict[str, str],
    dotted_lower: dict[str, str],
) -> list[ExtractedExpectation]:
    results: list[ExtractedExpectation] = []

    file_symbols = sorted(set(GET_STREAM_RE.findall(method.body)))
    resolved_files = [
        _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        for symbol in file_symbols
    ]
    resolved_files = [path for path in resolved_files if path is not None]
    single_file = resolved_files[0] if len(resolved_files) == 1 else None

    versions = list(VALIDATOR_VERSION_RE.findall(method.body))
    if VALIDATOR_DEFAULT_RE.search(method.body):
        versions.append("Office2007")
    normalized_versions = _normalize_versions(versions)
    has_empty_validate = "Assert.Empty(" in method.body and "Validate(" in method.body

    for match in INLINE_FILE_COUNT_RE.finditer(method.attributes):
        symbol = match.group(1)
        expected = int(match.group(2))
        relpath = _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        if relpath is None:
            continue
        results.append(
            _build_expectation(
                relpath=relpath,
                kind="inline_file_count",
                versions=normalized_versions,
                expected_count=expected,
                source=method,
            )
        )

    for match in INLINE_FILE_ONLY_RE.finditer(method.attributes):
        symbol = match.group(1)
        if symbol.startswith("FileFormatVersions."):
            continue
        relpath = _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        if relpath is None:
            continue
        if has_empty_validate and len(normalized_versions) == 1:
            results.append(
                _build_expectation(
                    relpath=relpath,
                    kind="inline_file_assert_empty",
                    versions=normalized_versions,
                    expected_count=0,
                    source=method,
                )
            )

    for match in INLINE_VERSION_COUNT_RE.finditer(method.attributes):
        version = match.group(1)
        expected = int(match.group(2))
        if single_file is None:
            continue
        results.append(
            _build_expectation(
                relpath=single_file,
                kind="inline_version_count",
                versions=(version,),
                expected_count=expected,
                source=method,
            )
        )

    for match in INLINE_VERSION_ONLY_RE.finditer(method.attributes):
        if single_file is None or not has_empty_validate:
            continue
        version = match.group(1)
        results.append(
            _build_expectation(
                relpath=single_file,
                kind="inline_version_assert_empty",
                versions=(version,),
                expected_count=0,
                source=method,
            )
        )

    if (
        has_empty_validate
        and single_file is not None
        and len(normalized_versions) == 1
    ):
        results.append(
            _build_expectation(
                relpath=single_file,
                kind="assert_empty_single_version",
                versions=normalized_versions,
                expected_count=0,
                source=method,
            )
        )

    allowed_match = ASSERT_TRUE_ALLOWED_COUNTS_RE.search(method.body)
    if (
        allowed_match is not None
        and single_file is not None
        and len(normalized_versions) >= 1
    ):
        a = int(allowed_match.group(1))
        b = int(allowed_match.group(2))
        results.append(
            _build_expectation(
                relpath=single_file,
                kind="assert_true_allowed_counts",
                versions=normalized_versions,
                expected_counts=tuple(sorted({a, b})),
                source=method,
            )
        )

    helper_match = XLSX_HELPER_CALL_RE.search(method.body)
    if helper_match is not None and single_file is not None:
        a = int(helper_match.group(1))
        b = int(helper_match.group(2))
        results.append(
            _build_expectation(
                relpath=single_file,
                kind="helper_allowed_counts",
                versions=("Office2007", "Office2010", "Office2013"),
                expected_counts=tuple(sorted({a, b})),
                source=method,
            )
        )

    return results


def _aggregate_expectations(items: list[ExtractedExpectation]) -> dict[str, list[dict[str, Any]]]:
    by_file: dict[str, dict[str, dict[str, Any]]] = defaultdict(dict)
    for item in items:
        key_payload = {
            "kind": item.kind,
            "validator_versions": list(item.validator_versions),
        }
        if item.expected_error_count is not None:
            key_payload["expected_error_count"] = item.expected_error_count
        if item.expected_error_counts:
            key_payload["expected_error_counts"] = list(item.expected_error_counts)
        key = json.dumps(key_payload, sort_keys=True)

        existing = by_file[item.relpath].get(key)
        if existing is None:
            by_file[item.relpath][key] = {
                **key_payload,
                "source_tests": [
                    {
                        "file": item.source_file,
                        "name": item.source_method,
                        "line": item.source_line,
                    }
                ],
            }
            continue

        source_list = existing["source_tests"]
        assert isinstance(source_list, list)
        source_list.append(
            {
                "file": item.source_file,
                "name": item.source_method,
                "line": item.source_line,
            }
        )

    finalized: dict[str, list[dict[str, Any]]] = {}
    for relpath, grouped in by_file.items():
        values = list(grouped.values())
        values.sort(
            key=lambda row: (
                str(row.get("kind", "")),
                ",".join(row.get("validator_versions", [])),
                int(row.get("expected_error_count", -1)),
            )
        )
        for value in values:
            tests = value["source_tests"]
            assert isinstance(tests, list)
            tests.sort(key=lambda row: (str(row.get("file", "")), int(row.get("line", 0))))
        finalized[relpath] = values
    return finalized


def _collect_expectations(
    sdk_root: Path,
    symbol_to_resource: dict[str, str],
    dotted_exact: dict[str, str],
    dotted_lower: dict[str, str],
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, int]]:
    all_items: list[ExtractedExpectation] = []
    stats = {
        "methods_scanned": 0,
        "raw_expectations": 0,
        "resolved_expectations": 0,
        "files_with_expectations": 0,
    }
    tests_root = sdk_root / TESTS_DIR
    for test_file in sorted(tests_root.rglob("*.cs")):
        methods = _iter_methods(test_file)
        stats["methods_scanned"] += len(methods)
        for method in methods:
            extracted = _extract_from_method(method, symbol_to_resource, dotted_exact, dotted_lower)
            stats["raw_expectations"] += len(extracted)
            all_items.extend(extracted)

    aggregated = _aggregate_expectations(all_items)
    stats["resolved_expectations"] = sum(len(values) for values in aggregated.values())
    stats["files_with_expectations"] = len(aggregated)
    return aggregated, stats


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract SDK expectations into corpus manifest.")
    parser.add_argument(
        "--sdk-root",
        type=Path,
        default=DEFAULT_SDK_ROOT,
        help=f"Path to Open XML SDK clone (default: {DEFAULT_SDK_ROOT})",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        default=DEFAULT_MANIFEST,
        help=f"Input manifest (default: {DEFAULT_MANIFEST})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output manifest path (default: overwrite --manifest)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Analyze and print stats without writing output.",
    )
    args = parser.parse_args()

    sdk_root = args.sdk_root.resolve()
    manifest_path = args.manifest.resolve()
    output_path = args.output.resolve() if args.output is not None else manifest_path

    if not manifest_path.exists():
        print(f"Manifest not found: {manifest_path}")
        return 2
    if not (sdk_root / ASSETS_DIR).exists():
        print(f"SDK assets directory not found: {sdk_root / ASSETS_DIR}")
        return 2
    if not (sdk_root / TESTS_DIR).exists():
        print(f"SDK tests directory not found: {sdk_root / TESTS_DIR}")
        return 2

    manifest = _load_manifest(manifest_path)
    dotted_exact, dotted_lower = _build_relpath_indexes(manifest)
    symbol_to_resource = _parse_constant_symbols(sdk_root / ASSETS_DIR)
    expectations_by_file, stats = _collect_expectations(
        sdk_root=sdk_root,
        symbol_to_resource=symbol_to_resource,
        dotted_exact=dotted_exact,
        dotted_lower=dotted_lower,
    )

    files = manifest.get("files", [])
    if not isinstance(files, list):
        print("Manifest has invalid 'files' payload.")
        return 2

    for entry in files:
        relpath = entry.get("source_relpath")
        if not isinstance(relpath, str):
            continue
        entry["expectations"] = expectations_by_file.get(relpath, [])

    manifest["expectation_stats"] = stats
    manifest["expectation_extraction"] = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "sdk_root": str(sdk_root),
        "script": "scripts/corpus/extract_sdk_expectations.py",
    }
    manifest["matched_expectation_files"] = stats["files_with_expectations"]

    print(f"SDK root: {sdk_root}")
    print(f"Manifest: {manifest_path}")
    print(f"Methods scanned: {stats['methods_scanned']}")
    print(f"Raw expectations: {stats['raw_expectations']}")
    print(f"Resolved expectations: {stats['resolved_expectations']}")
    print(f"Files with expectations: {stats['files_with_expectations']}")

    if args.dry_run:
        print("Dry run enabled, nothing written.")
        return 0

    output_path.parent.mkdir(parents=True, exist_ok=True)
    _save_manifest(output_path, manifest)
    print(f"Wrote manifest: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
