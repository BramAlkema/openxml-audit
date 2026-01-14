"""Compare openxml_audit output with the .NET Open XML SDK validator."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import OpenXmlValidator  # noqa: E402


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


def _local_runner(pptx_path: Path) -> list[dict]:
    project = ROOT / "scripts" / "sdk_compare" / "OpenXmlSdkValidator.csproj"
    cmd = ["dotnet", "run", "--project", str(project), "--", str(pptx_path)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            "dotnet validator failed:\n"
            f"stdout:\n{result.stdout}\n"
            f"stderr:\n{result.stderr}"
        )
    return json.loads(result.stdout) if result.stdout.strip() else []


def run_sdk_validator(pptx_path: Path, runner: str) -> list[dict]:
    if runner == "docker":
        raw = _docker_runner([pptx_path])
    else:
        raw = _local_runner(pptx_path)
    if not raw:
        return []
    if isinstance(raw, list) and isinstance(raw[0], dict) and "Errors" in raw[0]:
        return raw[0].get("Errors", [])
    return raw


def run_python_validator(pptx_path: Path) -> list[dict]:
    validator = OpenXmlValidator()
    result = validator.validate(pptx_path)
    return [
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


def normalize_error(error: dict) -> tuple[str, str, str]:
    return (
        (error.get("Description") or "").strip(),
        error.get("Part") or "",
        error.get("Path") or "",
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx_path", type=Path, help="Path to PPTX file")
    parser.add_argument(
        "--sdk-runner",
        choices=["auto", "local", "docker"],
        default="auto",
        help="How to run the SDK validator.",
    )
    parser.add_argument("--show-mismatches", action="store_true")
    parser.add_argument("--max-mismatches", type=int, default=20)
    args = parser.parse_args()

    runner = args.sdk_runner
    if runner == "auto":
        runner = "local" if shutil.which("dotnet") else "docker"
    if runner == "local" and not shutil.which("dotnet"):
        print("dotnet not found; install .NET SDK or use --sdk-runner docker.", file=sys.stderr)
        return 2

    if not args.pptx_path.exists():
        print(f"File not found: {args.pptx_path}", file=sys.stderr)
        return 2

    try:
        sdk_errors = run_sdk_validator(args.pptx_path, runner)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    py_errors = run_python_validator(args.pptx_path)

    sdk_set = {normalize_error(e) for e in sdk_errors}
    py_set = {normalize_error(e) for e in py_errors}

    only_sdk = sdk_set - py_set
    only_py = py_set - sdk_set

    print(f"SDK errors: {len(sdk_errors)}")
    print(f"Python errors: {len(py_errors)}")
    print(f"Only SDK: {len(only_sdk)}")
    print(f"Only Python: {len(only_py)}")

    if args.show_mismatches:
        if only_sdk:
            print("\nOnly SDK (first items):")
            for item in list(only_sdk)[: args.max_mismatches]:
                print(f"- {item}")
        if only_py:
            print("\nOnly Python (first items):")
            for item in list(only_py)[: args.max_mismatches]:
                print(f"- {item}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
