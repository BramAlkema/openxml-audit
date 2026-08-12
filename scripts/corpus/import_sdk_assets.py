#!/usr/bin/env python3
"""Import Open XML SDK test assets into a local corpus manifest."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SDK_ROOT = Path("/tmp/openxml-sdk-upstream")
DEFAULT_OUTPUT_ROOT = REPO_ROOT / "data" / "corpus" / "sdk_seed"
ASSETS_REL_PATH = Path("test/DocumentFormat.OpenXml.Tests.Assets/assets")
TEST_FILES_REL_PATH = Path("test/DocumentFormat.OpenXml.Tests.Assets/TestAssets.TestFiles.cs")
TESTS_REL_PATH = Path("test/DocumentFormat.OpenXml.Tests")
DEFAULT_EXTENSIONS = {
    ".docx",
    ".docm",
    ".dotx",
    ".dotm",
    ".pptx",
    ".pptm",
    ".potx",
    ".potm",
    ".ppsx",
    ".ppsm",
    ".ppam",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
}

METHOD_SIGNATURE_RE = re.compile(
    r"^\s*public\s+(?:static\s+)?(?:async\s+)?(?:void|Task(?:<[^>]+>)?)\s+([A-Za-z0-9_]+)\s*\("
)
INLINE_TESTFILE_RE = re.compile(r"\[InlineData\(\s*TestFiles\.([A-Za-z0-9_]+)\s*,\s*(-?\d+)\s*\)\]")
INLINE_VERSION_RE = re.compile(
    r"\[InlineData\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*,\s*(-?\d+)\s*\)\]"
)
OPENXML_VALIDATOR_RE = re.compile(r"OpenXmlValidator\(FileFormatVersions\.([A-Za-z0-9_]+)\)")
GET_STREAM_RE = re.compile(r"GetStream\(TestFiles\.([A-Za-z0-9_]+)")
CONST_RE = re.compile(r'public const string ([A-Za-z0-9_]+)\s*=\s*"([^"]+)";')
XLSX_HELPER_CALL_RE = re.compile(
    r"XlsxValidationHelper\(\s*stream\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)"
)


@dataclass(frozen=True)
class MethodBlock:
    """Container for a parsed C# method block."""

    name: str
    start_line: int
    attributes: str
    body: str
    source: str


def _normalize_extensions(raw: str | None) -> set[str]:
    if raw is None:
        return set(DEFAULT_EXTENSIONS)
    values = {part.strip().lower() for part in raw.split(",") if part.strip()}
    return {f".{value}" if not value.startswith(".") else value for value in values}


def _resource_to_relative_path(resource: str) -> Path:
    """Convert SDK resource naming to a relative asset path."""
    if "." not in resource:
        return Path(resource)
    prefix, filename = resource.split(".", 1)
    return Path(prefix) / filename


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_test_file_constants(sdk_root: Path) -> dict[str, str]:
    const_file = sdk_root / TEST_FILES_REL_PATH
    if not const_file.exists():
        raise FileNotFoundError(f"Missing expected SDK constants file: {const_file}")
    constants: dict[str, str] = {}
    text = const_file.read_text(encoding="utf-8")
    for match in CONST_RE.finditer(text):
        key = match.group(1)
        value = match.group(2)
        if value.startswith("TestFiles."):
            constants[key] = value
    return constants


def _find_method_end(lines: list[str], start_index: int) -> int:
    open_line = -1
    open_col = -1
    for line_index in range(start_index, len(lines)):
        col = lines[line_index].find("{")
        if col >= 0:
            open_line = line_index
            open_col = col
            break
    if open_line < 0:
        return start_index

    depth = 0
    for line_index in range(open_line, len(lines)):
        start_col = open_col if line_index == open_line else 0
        line = lines[line_index]
        for col in range(start_col, len(line)):
            char = line[col]
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    return line_index
    return len(lines) - 1


def _collect_attributes(lines: list[str], method_start: int) -> str:
    attrs: list[str] = []
    line_index = method_start - 1
    while line_index >= 0:
        stripped = lines[line_index].strip()
        if stripped.startswith("["):
            attrs.append(lines[line_index])
            line_index -= 1
            continue
        if not stripped:
            line_index -= 1
            continue
        break
    attrs.reverse()
    return "\n".join(attrs)


def _iter_method_blocks(path: Path) -> list[MethodBlock]:
    lines = path.read_text(encoding="utf-8").splitlines()
    methods: list[MethodBlock] = []
    line_index = 0
    while line_index < len(lines):
        match = METHOD_SIGNATURE_RE.match(lines[line_index])
        if match is None:
            line_index += 1
            continue
        method_name = match.group(1)
        end_line = _find_method_end(lines, line_index)
        methods.append(
            MethodBlock(
                name=method_name,
                start_line=line_index + 1,
                attributes=_collect_attributes(lines, line_index),
                body="\n".join(lines[line_index : end_line + 1]),
                source=str(path),
            )
        )
        line_index = end_line + 1
    return methods


def _extract_expectations(
    sdk_root: Path, constants: dict[str, str]
) -> tuple[dict[str, list[dict[str, object]]], dict[str, int]]:
    expected_by_relpath: dict[str, list[dict[str, object]]] = defaultdict(list)
    stats = {
        "methods_scanned": 0,
        "expectations_found": 0,
        "missing_constants": 0,
        "ambiguous_versioned_methods": 0,
        "assert_empty_methods": 0,
        "helper_methods": 0,
    }

    tests_root = sdk_root / TESTS_REL_PATH
    if not tests_root.exists():
        return {}, stats

    for test_file in sorted(tests_root.rglob("*.cs")):
        for method in _iter_method_blocks(test_file):
            stats["methods_scanned"] += 1
            versions = sorted(set(OPENXML_VALIDATOR_RE.findall(method.body)))
            stream_constants = sorted(set(GET_STREAM_RE.findall(method.body)))

            for inline in INLINE_TESTFILE_RE.finditer(method.attributes):
                const_name, expected_count = inline.group(1), int(inline.group(2))
                resource = constants.get(const_name)
                if resource is None:
                    stats["missing_constants"] += 1
                    continue
                rel_path = _resource_to_relative_path(resource).as_posix()
                expected_by_relpath[rel_path].append(
                    {
                        "kind": "inline_testfile_count",
                        "source_test": f"{test_file.name}::{method.name}",
                        "source_line": method.start_line,
                        "expected_error_count": expected_count,
                        "validator_versions": versions,
                    }
                )
                stats["expectations_found"] += 1

            for inline in INLINE_VERSION_RE.finditer(method.attributes):
                version, expected_count = inline.group(1), int(inline.group(2))
                if len(stream_constants) != 1:
                    stats["ambiguous_versioned_methods"] += 1
                    continue
                stream_const = stream_constants[0]
                resource = constants.get(stream_const)
                if resource is None:
                    stats["missing_constants"] += 1
                    continue
                rel_path = _resource_to_relative_path(resource).as_posix()
                expected_by_relpath[rel_path].append(
                    {
                        "kind": "inline_version_count",
                        "source_test": f"{test_file.name}::{method.name}",
                        "source_line": method.start_line,
                        "file_format_version": version,
                        "expected_error_count": expected_count,
                    }
                )
                stats["expectations_found"] += 1

            if (
                "Assert.Empty(" in method.body
                and "Validate(" in method.body
                and len(stream_constants) == 1
                and len(versions) == 1
            ):
                stream_const = stream_constants[0]
                resource = constants.get(stream_const)
                if resource is None:
                    stats["missing_constants"] += 1
                else:
                    rel_path = _resource_to_relative_path(resource).as_posix()
                    expected_by_relpath[rel_path].append(
                        {
                            "kind": "assert_empty_single_version",
                            "source_test": f"{test_file.name}::{method.name}",
                            "source_line": method.start_line,
                            "file_format_version": versions[0],
                            "expected_error_count": 0,
                        }
                    )
                    stats["expectations_found"] += 1
                    stats["assert_empty_methods"] += 1

            helper_match = XLSX_HELPER_CALL_RE.search(method.body)
            if helper_match and len(stream_constants) == 1:
                stream_const = stream_constants[0]
                resource = constants.get(stream_const)
                if resource is None:
                    stats["missing_constants"] += 1
                else:
                    expected_a = int(helper_match.group(1))
                    expected_b = int(helper_match.group(2))
                    rel_path = _resource_to_relative_path(resource).as_posix()
                    expected_by_relpath[rel_path].append(
                        {
                            "kind": "helper_allowed_counts",
                            "source_test": f"{test_file.name}::{method.name}",
                            "source_line": method.start_line,
                            "expected_error_counts": sorted({expected_a, expected_b}),
                            "notes": "From XlsxValidationHelper(stream, a, b) invocation.",
                        }
                    )
                    stats["expectations_found"] += 1
                    stats["helper_methods"] += 1

    for rel_path, entries in expected_by_relpath.items():
        grouped: dict[str, dict[str, object]] = {}
        for entry in entries:
            source_test = str(entry["source_test"])
            source_line = int(entry["source_line"])
            normalized = {k: v for k, v in entry.items() if k not in {"source_test", "source_line"}}
            key = json.dumps(normalized, sort_keys=True)
            if key not in grouped:
                grouped[key] = {
                    **normalized,
                    "source_tests": [{"name": source_test, "line": source_line}],
                }
                continue
            source_tests = grouped[key]["source_tests"]
            assert isinstance(source_tests, list)
            source_tests.append({"name": source_test, "line": source_line})
        expected_by_relpath[rel_path] = sorted(
            grouped.values(),
            key=lambda item: (
                str(item.get("kind", "")),
                str(item.get("file_format_version", "")),
                int(item.get("expected_error_count", 0))
                if isinstance(item.get("expected_error_count"), int)
                else 0,
            ),
        )

    return dict(expected_by_relpath), stats


def _collect_asset_files(
    assets_root: Path, include_dirs: list[str], extensions: set[str], max_files: int
) -> list[Path]:
    files: set[Path] = set()
    search_roots: list[Path]
    if include_dirs:
        search_roots = [assets_root / include_dir for include_dir in include_dirs]
    else:
        search_roots = [assets_root]

    for root in search_roots:
        if not root.exists():
            raise FileNotFoundError(f"Include directory not found: {root}")
        if root.is_file():
            if root.suffix.lower() in extensions:
                files.add(root)
            continue
        for file_path in root.rglob("*"):
            if file_path.is_file() and file_path.suffix.lower() in extensions:
                files.add(file_path)

    selected = sorted(files)
    if max_files > 0:
        selected = selected[:max_files]
    return selected


def _build_manifest(
    sdk_root: Path,
    assets_root: Path,
    selected_files: list[Path],
    output_root: Path,
    copy_files: bool,
    expectations: dict[str, list[dict[str, object]]],
) -> dict[str, object]:
    files_payload: list[dict[str, object]] = []
    copied_count = 0
    for source_path in selected_files:
        rel_path = source_path.relative_to(assets_root)
        rel_posix = rel_path.as_posix()
        destination_path = output_root / "files" / rel_path
        if copy_files:
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, destination_path)
            copied_count += 1
        stat = source_path.stat()
        files_payload.append(
            {
                "source_relpath": rel_posix,
                "destination_relpath": f"files/{rel_posix}" if copy_files else None,
                "size_bytes": stat.st_size,
                "sha256": _sha256_file(source_path),
                "extension": source_path.suffix.lower(),
                "expectations": expectations.get(rel_posix, []),
            }
        )

    return {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "sdk_root": str(sdk_root),
        "assets_root": str(assets_root),
        "output_root": str(output_root),
        "copied_files": copied_count,
        "total_files": len(files_payload),
        "files": files_payload,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Import Open XML SDK assets and generate a corpus manifest."
    )
    parser.add_argument(
        "--sdk-root",
        type=Path,
        default=DEFAULT_SDK_ROOT,
        help=f"Path to Open XML SDK clone (default: {DEFAULT_SDK_ROOT})",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT,
        help=f"Corpus output root (default: {DEFAULT_OUTPUT_ROOT})",
    )
    parser.add_argument(
        "--include-dir",
        action="append",
        default=[],
        help=(
            "Directory under SDK assets root to include (repeatable). Default: include all assets."
        ),
    )
    parser.add_argument(
        "--extensions",
        default=None,
        help=(
            "Comma-separated extensions to include, e.g. docx,pptx,xlsx. "
            "Default: common OOXML package extensions."
        ),
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Maximum files to include (0 means no limit).",
    )
    parser.add_argument(
        "--manifest-only",
        action="store_true",
        help="Do not copy file payloads; write only manifest metadata.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not write files; print summary only.",
    )
    args = parser.parse_args()

    sdk_root = args.sdk_root.resolve()
    assets_root = sdk_root / ASSETS_REL_PATH
    if not assets_root.exists():
        print(f"Assets root not found: {assets_root}")
        return 2

    extensions = _normalize_extensions(args.extensions)
    selected_files = _collect_asset_files(
        assets_root=assets_root,
        include_dirs=args.include_dir,
        extensions=extensions,
        max_files=args.max_files,
    )
    if not selected_files:
        print("No matching files selected.")
        return 2

    constants = _load_test_file_constants(sdk_root)
    expectations, expectation_stats = _extract_expectations(sdk_root, constants)

    output_root = args.output_root.resolve()
    copy_files = not args.manifest_only
    if args.dry_run:
        copy_files = False

    manifest = _build_manifest(
        sdk_root=sdk_root,
        assets_root=assets_root,
        selected_files=selected_files,
        output_root=output_root,
        copy_files=copy_files,
        expectations=expectations,
    )
    manifest["expectation_stats"] = expectation_stats
    manifest["settings"] = {
        "include_dirs": args.include_dir,
        "extensions": sorted(extensions),
        "max_files": args.max_files,
        "manifest_only": args.manifest_only,
        "dry_run": args.dry_run,
    }
    matched_expectation_files = sum(
        1
        for file_entry in manifest["files"]
        if file_entry["expectations"]  # type: ignore[index]
    )
    manifest["matched_expectation_files"] = matched_expectation_files

    print(f"SDK root: {sdk_root}")
    print(f"Assets root: {assets_root}")
    print(f"Selected files: {manifest['total_files']}")
    print(f"Files with expectations: {matched_expectation_files}")
    print(f"Extracted expectations: {expectation_stats['expectations_found']}")

    if args.dry_run:
        print("Dry run enabled, nothing written.")
        return 0

    output_root.mkdir(parents=True, exist_ok=True)
    manifest_path = output_root / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(f"Wrote manifest: {manifest_path}")
    if copy_files:
        print(f"Copied files into: {output_root / 'files'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
