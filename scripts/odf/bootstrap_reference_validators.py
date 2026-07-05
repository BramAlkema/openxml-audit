#!/usr/bin/env python3
"""Bootstrap external ODF reference validators and emit command templates."""

from __future__ import annotations

import argparse
import json
import shlex
import shutil
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_WORK_ROOT = Path("/tmp/openxml_audit_odf_reference_tools")
DEFAULT_ODF_TOOLKIT_REPO = "https://github.com/tdf/odftoolkit.git"
DEFAULT_ODF_TOOLKIT_REF = "v0.13.0"
DEFAULT_OPF_REPO = "https://github.com/opf-labs/odf-validator.git"
DEFAULT_OPF_FALLBACK_REPO = "https://github.com/openpreserve/odf-validator.git"
DEFAULT_OPF_REF = "v0.20-alpha-2"
DEFAULT_MAVEN_IMAGE = "maven:3.9.9-eclipse-temurin-17"
DEFAULT_RUNTIME_IMAGE = "eclipse-temurin:17-jre"


def _require_executable(name: str) -> None:
    if shutil.which(name) is None:
        raise RuntimeError(
            f"Required executable '{name}' was not found in PATH. "
            "Install it and retry."
        )


def _run_checked(command: list[str], *, cwd: Path | None = None) -> str:
    completed = subprocess.run(
        command,
        cwd=str(cwd) if cwd else None,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "Command failed with exit code "
            f"{completed.returncode}: {shlex.join(command)}\n"
            f"stdout:\n{completed.stdout}\n"
            f"stderr:\n{completed.stderr}"
        )
    return completed.stdout


def _java_runtime_ready() -> bool:
    java = shutil.which("java")
    if java is None:
        return False
    completed = subprocess.run(
        [java, "-version"],
        capture_output=True,
        text=True,
        check=False,
    )
    return completed.returncode == 0


def _resolve_maven_mode(requested: str) -> str:
    if requested == "local":
        _require_executable("mvn")
        return "local"
    if requested == "docker":
        _require_executable("docker")
        return "docker"
    if shutil.which("mvn") is not None:
        return "local"
    if shutil.which("docker") is not None:
        return "docker"
    raise RuntimeError(
        "No Maven build path available. Install Maven (mvn) or Docker, "
        "or pass --maven-mode with an available backend."
    )


def _resolve_runtime_mode(requested: str) -> str:
    if requested == "local":
        if not _java_runtime_ready():
            raise RuntimeError(
                "Local runtime mode requires a working Java runtime. "
                "Install Java or use --runtime-mode docker."
            )
        return "local"
    if requested == "docker":
        _require_executable("docker")
        return "docker"

    if _java_runtime_ready():
        return "local"
    if shutil.which("docker") is not None:
        return "docker"

    raise RuntimeError(
        "No runtime backend available for executing validators. "
        "Install Java or Docker, or pass --runtime-mode with an available backend."
    )


def _run_maven(
    args: list[str],
    *,
    cwd: Path,
    mode: str,
    maven_image: str,
) -> None:
    if mode == "local":
        _run_checked(["mvn", *args], cwd=cwd)
        return
    _run_checked(
        [
            "docker",
            "run",
            "--rm",
            "-v",
            f"{cwd.resolve()}:/work",
            "-w",
            "/work",
            maven_image,
            "mvn",
            *args,
        ]
    )


def _clone_repo(repo: str, ref: str, target_dir: Path) -> None:
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.parent.mkdir(parents=True, exist_ok=True)
    _run_checked(
        ["git", "clone", "--depth", "1", "--branch", ref, repo, str(target_dir)]
    )


def _select_odf_toolkit_jar(toolkit_root: Path) -> Path:
    target_dir = toolkit_root / "validator" / "target"
    if not target_dir.exists():
        raise FileNotFoundError(f"ODF Toolkit target directory not found: {target_dir}")

    candidates = sorted(target_dir.glob("odfvalidator-*-jar-with-dependencies.jar"))
    if not candidates:
        candidates = sorted(
            path
            for path in target_dir.glob("odfvalidator-*.jar")
            if "javadoc" not in path.name
            and "sources" not in path.name
            and "original-" not in path.name
        )
    if not candidates:
        raise FileNotFoundError(
            "Unable to locate ODF Toolkit validator jar in "
            f"{target_dir}"
        )
    return candidates[-1]


def _select_opf_launcher(opf_root: Path) -> tuple[str, Path]:
    script_candidates = (
        opf_root / "odf-validator",
        opf_root / "target" / "odf-validator",
        opf_root / "odf-validator.sh",
    )
    for candidate in script_candidates:
        if candidate.exists() and candidate.is_file():
            candidate.chmod(candidate.stat().st_mode | 0o111)
            return ("script", candidate)

    jar_candidates = sorted(
        path
        for path in (opf_root / "target").glob("*.jar")
        if "javadoc" not in path.name and "sources" not in path.name
    )
    if jar_candidates:
        return ("jar", jar_candidates[-1])

    raise FileNotFoundError(
        "Unable to locate OPF validator launcher script or jar in "
        f"{opf_root}"
    )


def _join_shell_tokens(tokens: list[str]) -> str:
    return " ".join(shlex.quote(token) for token in tokens)


def _container_tool_path(target_root: Path, inner_path: Path) -> str:
    relative = inner_path.resolve().relative_to(target_root.resolve())
    return f"/tool/{relative.as_posix()}"


def _build_odf_toolkit_command_template(
    *,
    target: Path,
    jar_path: Path,
    runtime_mode: str,
    runtime_image: str,
) -> str:
    if runtime_mode == "local":
        return _join_shell_tokens(["java", "-jar", str(jar_path), "{file}"])

    return _join_shell_tokens(
        [
            "docker",
            "run",
            "--rm",
            "-v",
            f"{target.resolve()}:/tool:ro",
            "-v",
            "{file_dir}:/input:ro",
            runtime_image,
            "java",
            "-jar",
            _container_tool_path(target, jar_path),
            "/input/{file_name}",
        ]
    )


def _build_opf_command_template(
    *,
    target: Path,
    launcher_type: str,
    launcher_path: Path,
    runtime_mode: str,
    runtime_image: str,
) -> str:
    if runtime_mode == "local":
        if launcher_type == "script":
            return _join_shell_tokens([str(launcher_path), "-o", "JSON", "{file}"])
        return _join_shell_tokens(
            ["java", "-jar", str(launcher_path), "-o", "JSON", "{file}"]
        )

    container_launcher = _container_tool_path(target, launcher_path)
    tokens = [
        "docker",
        "run",
        "--rm",
        "-v",
        f"{target.resolve()}:/tool:ro",
        "-v",
        "{file_dir}:/input:ro",
    ]
    if launcher_type == "script":
        tokens.extend(
            [
                "-w",
                "/tool",
                runtime_image,
                container_launcher,
                "-o",
                "JSON",
                "/input/{file_name}",
            ]
        )
    else:
        tokens.extend(
            [
                runtime_image,
                "java",
                "-jar",
                container_launcher,
                "-o",
                "JSON",
                "/input/{file_name}",
            ]
        )
    return _join_shell_tokens(tokens)


def _build_odf_toolkit(
    *,
    repo: str,
    ref: str,
    root: Path,
    maven_mode: str,
    maven_image: str,
    runtime_mode: str,
    runtime_image: str,
) -> dict[str, Any]:
    target = root / "odf_toolkit"
    _clone_repo(repo, ref, target)
    _run_maven(
        ["-q", "-DskipTests", "-pl", "validator", "-am", "package"],
        cwd=target,
        mode=maven_mode,
        maven_image=maven_image,
    )
    jar_path = _select_odf_toolkit_jar(target)
    command_template = _build_odf_toolkit_command_template(
        target=target,
        jar_path=jar_path,
        runtime_mode=runtime_mode,
        runtime_image=runtime_image,
    )
    return {
        "repo": repo,
        "ref": ref,
        "checkout_path": str(target),
        "runtime_mode": runtime_mode,
        "command_template": command_template,
        # The toolkit command always runs `java -jar <absolute path>`, so it is
        # cwd-independent; no working directory needs to be pinned.
        "working_dir": "",
    }


def _build_opf_validator(
    *,
    repo: str,
    fallback_repo: str | None,
    ref: str,
    root: Path,
    maven_mode: str,
    maven_image: str,
    runtime_mode: str,
    runtime_image: str,
) -> dict[str, Any]:
    target = root / "opf"
    selected_repo = repo
    try:
        _clone_repo(repo, ref, target)
    except RuntimeError:
        if fallback_repo is None:
            raise
        selected_repo = fallback_repo
        _clone_repo(fallback_repo, ref, target)

    _run_maven(
        ["-q", "-DskipTests", "package"],
        cwd=target,
        mode=maven_mode,
        maven_image=maven_image,
    )
    launcher_type, launcher_path = _select_opf_launcher(target)
    command_template = _build_opf_command_template(
        target=target,
        launcher_type=launcher_type,
        launcher_path=launcher_path,
        runtime_mode=runtime_mode,
        runtime_image=runtime_image,
    )
    # The OPF launcher script resolves its jar via a repo-relative path
    # (e.g. `odf-apps/target/odf-apps-<ver>-jar-with-dependencies.jar`), so it
    # only works when executed from the checkout root. Docker mode handles this
    # with `-w /tool`; local script mode must set the process working directory
    # explicitly or every invocation fails with "Unable to access jarfile".
    working_dir = (
        str(target) if runtime_mode == "local" and launcher_type == "script" else ""
    )
    return {
        "repo": selected_repo,
        "ref": ref,
        "checkout_path": str(target),
        "launcher_type": launcher_type,
        "runtime_mode": runtime_mode,
        "command_template": command_template,
        "working_dir": working_dir,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build ODF reference validators and emit command templates."
    )
    parser.add_argument(
        "--work-root",
        type=Path,
        default=DEFAULT_WORK_ROOT,
        help=f"Build workspace root (default: {DEFAULT_WORK_ROOT})",
    )
    parser.add_argument(
        "--odf-toolkit-repo",
        type=str,
        default=DEFAULT_ODF_TOOLKIT_REPO,
    )
    parser.add_argument(
        "--odf-toolkit-ref",
        type=str,
        default=DEFAULT_ODF_TOOLKIT_REF,
    )
    parser.add_argument(
        "--opf-repo",
        type=str,
        default=DEFAULT_OPF_REPO,
    )
    parser.add_argument(
        "--opf-fallback-repo",
        type=str,
        default=DEFAULT_OPF_FALLBACK_REPO,
    )
    parser.add_argument(
        "--opf-ref",
        type=str,
        default=DEFAULT_OPF_REF,
    )
    parser.add_argument(
        "--maven-mode",
        type=str,
        choices=("auto", "local", "docker"),
        default="auto",
        help="Maven backend selection: auto (default), local, or docker.",
    )
    parser.add_argument(
        "--maven-image",
        type=str,
        default=DEFAULT_MAVEN_IMAGE,
        help=f"Docker Maven image when --maven-mode=docker (default: {DEFAULT_MAVEN_IMAGE})",
    )
    parser.add_argument(
        "--runtime-mode",
        type=str,
        choices=("auto", "local", "docker"),
        default="auto",
        help="Runtime backend for executing validator commands: auto, local, or docker.",
    )
    parser.add_argument(
        "--runtime-image",
        type=str,
        default=DEFAULT_RUNTIME_IMAGE,
        help=(
            "Docker runtime image when --runtime-mode=docker "
            f"(default: {DEFAULT_RUNTIME_IMAGE})"
        ),
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        default=None,
        help="Optional JSON file path for the generated command templates.",
    )
    parser.add_argument(
        "--print-shell",
        action="store_true",
        help="Print export-friendly shell lines in addition to JSON.",
    )
    args = parser.parse_args()

    _require_executable("git")
    maven_mode = _resolve_maven_mode(args.maven_mode)
    runtime_mode = _resolve_runtime_mode(args.runtime_mode)

    work_root = args.work_root.resolve()
    if work_root.exists():
        shutil.rmtree(work_root)
    work_root.mkdir(parents=True, exist_ok=True)

    odf_toolkit = _build_odf_toolkit(
        repo=args.odf_toolkit_repo,
        ref=args.odf_toolkit_ref,
        root=work_root,
        maven_mode=maven_mode,
        maven_image=args.maven_image,
        runtime_mode=runtime_mode,
        runtime_image=args.runtime_image,
    )
    opf = _build_opf_validator(
        repo=args.opf_repo,
        fallback_repo=args.opf_fallback_repo or None,
        ref=args.opf_ref,
        root=work_root,
        maven_mode=maven_mode,
        maven_image=args.maven_image,
        runtime_mode=runtime_mode,
        runtime_image=args.runtime_image,
    )

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "workspace": str(work_root),
        "maven_mode": maven_mode,
        "runtime_mode": runtime_mode,
        "runtime_image": args.runtime_image if runtime_mode == "docker" else None,
        "odf_toolkit": odf_toolkit,
        "opf": opf,
        "odf_toolkit_cmd": odf_toolkit["command_template"],
        "opf_cmd": opf["command_template"],
        "odf_toolkit_working_dir": odf_toolkit["working_dir"],
        "opf_working_dir": opf["working_dir"],
    }

    payload = json.dumps(output, indent=2)
    print(payload)
    if args.output_json is not None:
        output_path = args.output_json.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(payload, encoding="utf-8")
        print(f"Wrote bootstrap metadata: {output_path}")

    if args.print_shell:
        print(f"ODF_TOOLKIT_CMD={output['odf_toolkit_cmd']}")
        print(f"OPF_ODF_VALIDATOR_CMD={output['opf_cmd']}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
