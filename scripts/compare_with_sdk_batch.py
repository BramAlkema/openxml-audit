"""Batch compare openxml_audit output with the .NET Open XML SDK validator."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import FileFormat, OpenXmlValidator  # noqa: E402

OOXML_EXTENSIONS = {
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
    ".thmx",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".xlam",
}


def _chunked(items: list[Path], chunk_size: int) -> Iterable[list[Path]]:
    for i in range(0, len(items), chunk_size):
        yield items[i : i + chunk_size]


def _collect_files(paths: list[Path], recursive: bool) -> list[Path]:
    files: set[Path] = set()
    for path in paths:
        if path.is_dir():
            if not recursive:
                raise ValueError(f"{path} is a directory. Use --recursive to include it.")
            for ext in OOXML_EXTENSIONS:
                files.update(path.rglob(f"*{ext}"))
        else:
            if path.suffix.lower() in OOXML_EXTENSIONS:
                files.add(path)
    return sorted(files)


def _docker_runner(paths: list[Path]) -> list[dict]:
    if not paths:
        return []
    host_root = ROOT.parent
    container_root = Path("/work")
    container_cwd = container_root / ROOT.relative_to(host_root)
    container_paths = []
    for path in paths:
        resolved = path.resolve()
        try:
            rel = resolved.relative_to(host_root)
        except ValueError as exc:
            raise RuntimeError(
                f"Path {resolved} is not under {host_root}; cannot mount into docker."
            ) from exc
        container_paths.append(str(container_root / rel))

    project = container_cwd / "scripts" / "sdk_compare" / "OpenXmlSdkValidator.csproj"
    cmd = [
        "docker",
        "run",
        "--rm",
        "-v",
        f"{host_root}:{container_root}",
        "-w",
        str(container_cwd),
        "mcr.microsoft.com/dotnet/sdk:8.0",
        "dotnet",
        "run",
        "--project",
        str(project),
        "--",
    ] + container_paths
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            "docker validator failed:\n"
            f"stdout:\n{result.stdout}\n"
            f"stderr:\n{result.stderr}"
        )
    return json.loads(result.stdout) if result.stdout.strip() else []


def _local_runner(paths: list[Path]) -> list[dict]:
    project = ROOT / "scripts" / "sdk_compare" / "OpenXmlSdkValidator.csproj"
    cmd = ["dotnet", "run", "--project", str(project), "--"] + [str(p) for p in paths]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            "dotnet validator failed:\n"
            f"stdout:\n{result.stdout}\n"
            f"stderr:\n{result.stderr}"
        )
    return json.loads(result.stdout) if result.stdout.strip() else []


def _normalize_error(error: dict) -> tuple[str, str, str]:
    return (
        (error.get("Description") or "").strip(),
        error.get("Part") or "",
        error.get("Path") or "",
    )


def _convert_sdk_results(raw: list[dict]) -> dict[Path, list[dict]]:
    sdk_by_file: dict[Path, list[dict]] = {}
    if not raw:
        return sdk_by_file
    host_root = ROOT.parent
    container_root = Path("/work")
    for entry in raw:
        file_path = Path(entry.get("File", ""))
        if file_path.is_absolute() and str(file_path).startswith(str(container_root)):
            file_path = host_root / file_path.relative_to(container_root)
        sdk_by_file[file_path] = entry.get("Errors", [])
    return sdk_by_file


def _run_sdk(paths: list[Path], runner: str, chunk_size: int) -> dict[Path, list[dict]]:
    sdk_results: dict[Path, list[dict]] = {}
    for chunk in _chunked(paths, chunk_size):
        raw = _docker_runner(chunk) if runner == "docker" else _local_runner(chunk)
        sdk_results.update(_convert_sdk_results(raw))
    return sdk_results


def _run_python(paths: list[Path]) -> dict[Path, list[dict]]:
    validator = OpenXmlValidator(file_format=FileFormat.OFFICE_2019, max_errors=0)
    results: dict[Path, list[dict]] = {}
    for path in paths:
        result = validator.validate(path)
        results[path] = [
            {
                "Description": error.description,
                "Part": error.part_uri,
                "Path": error.path,
                "Node": error.node or "",
                "RelatedNode": error.related_node or "",
                "ErrorType": error.error_type.value if error.error_type else "",
            }
            for error in result.errors
        ]
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("paths", nargs="+", type=Path, help="Files or directories to compare")
    parser.add_argument(
        "--recursive",
        "-r",
        action="store_true",
        help="Recurse into directories",
    )
    parser.add_argument(
        "--sdk-runner",
        choices=["auto", "local", "docker"],
        default="docker",
        help="How to run the SDK validator.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=40,
        help="Number of files per SDK invocation.",
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Maximum number of files to compare (0 for no limit).",
    )
    parser.add_argument("--show-mismatches", action="store_true")
    parser.add_argument("--max-mismatches", type=int, default=10)
    args = parser.parse_args()

    try:
        files = _collect_files(args.paths, args.recursive)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    if args.max_files:
        files = files[: args.max_files]

    if not files:
        print("No matching OOXML files found.", file=sys.stderr)
        return 2

    runner = args.sdk_runner
    if runner == "auto":
        runner = "local" if shutil.which("dotnet") else "docker"
    if runner == "local" and not shutil.which("dotnet"):
        print("dotnet not found; install .NET SDK or use --sdk-runner docker.", file=sys.stderr)
        return 2

    sdk_results = _run_sdk(files, runner, args.chunk_size)
    py_results = _run_python(files)

    total_sdk = 0
    total_py = 0
    total_only_sdk = 0
    total_only_py = 0
    mismatches: list[tuple[Path, set[tuple[str, str, str]], set[tuple[str, str, str]]]] = []
    sdk_failures: list[tuple[Path, str]] = []

    for path in files:
        sdk_errors = sdk_results.get(path, [])
        py_errors = py_results.get(path, [])

        sdk_error_types = {e.get("ErrorType") for e in sdk_errors}
        if "Exception" in sdk_error_types or "FileNotFound" in sdk_error_types:
            for error in sdk_errors:
                sdk_failures.append((path, error.get("Description", "SDK failure")))
            continue

        sdk_set = {_normalize_error(e) for e in sdk_errors}
        py_set = {_normalize_error(e) for e in py_errors}

        total_sdk += len(sdk_errors)
        total_py += len(py_errors)

        only_sdk = sdk_set - py_set
        only_py = py_set - sdk_set
        total_only_sdk += len(only_sdk)
        total_only_py += len(only_py)

        if only_sdk or only_py:
            mismatches.append((path, only_sdk, only_py))

    print(f"Files compared: {len(files)}")
    print(f"SDK errors: {total_sdk}")
    print(f"Python errors: {total_py}")
    print(f"Only SDK: {total_only_sdk}")
    print(f"Only Python: {total_only_py}")
    print(f"Files with deltas: {len(mismatches)}")
    if sdk_failures:
        print(f"SDK failures: {len(sdk_failures)}")

    if args.show_mismatches and mismatches:
        for path, only_sdk, only_py in mismatches[: args.max_mismatches]:
            print(f"\n{path}")
            if only_sdk:
                print("  Only SDK:")
                for item in list(only_sdk)[: args.max_mismatches]:
                    print(f"    - {item}")
            if only_py:
                print("  Only Python:")
                for item in list(only_py)[: args.max_mismatches]:
                    print(f"    - {item}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
