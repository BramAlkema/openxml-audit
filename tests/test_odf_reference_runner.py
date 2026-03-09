from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path
from types import ModuleType

SCRIPT_PATH = (
    Path(__file__).resolve().parents[1]
    / "scripts"
    / "odf"
    / "run_reference_validators.py"
)


def _load_runner_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location("odf_reference_runner", SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load runner script module from {SCRIPT_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_format_command_expands_path_placeholders(tmp_path: Path) -> None:
    module = _load_runner_module()
    file_path = tmp_path / "sample.odt"
    file_path.write_text("stub", encoding="utf-8")

    command = module._format_command(
        ["runner", "--dir", "{file_dir}", "--name", "{file_name}"],
        file_path,
    )
    assert command == [
        "runner",
        "--dir",
        str(file_path.parent),
        "--name",
        file_path.name,
    ]


def test_run_reference_runner_marks_missing_java_runtime_unavailable(
    tmp_path: Path, monkeypatch
) -> None:
    module = _load_runner_module()
    sample = tmp_path / "sample.odt"
    sample.write_text("stub", encoding="utf-8")

    def fake_run(*args, **kwargs):  # type: ignore[no-untyped-def]
        return subprocess.CompletedProcess(
            args=args[0],
            returncode=1,
            stdout="",
            stderr=(
                "The operation couldn’t be completed. Unable to locate a Java Runtime.\n"
                "Please visit http://www.java.com for information on installing Java.\n"
            ),
        )

    monkeypatch.setattr(module.subprocess, "run", fake_run)

    config = module.RunnerConfig(
        name="odf_toolkit",
        command_template=["java", "-jar", "tool.jar", "{file}"],
    )
    result = module._run_reference_runner(
        config,
        sample,
        timeout_seconds=10,
    )

    assert result["status"] == "unavailable"
    assert "runtime dependency unavailable" in str(result.get("reason", "")).lower()
    assert result["issue_count"] == 0
    assert result["issues"] == []


def test_run_reference_runner_keeps_validation_findings_for_nonzero_exit(
    tmp_path: Path, monkeypatch
) -> None:
    module = _load_runner_module()
    sample = tmp_path / "sample.odt"
    sample.write_text("stub", encoding="utf-8")

    def fake_run(*args, **kwargs):  # type: ignore[no-untyped-def]
        return subprocess.CompletedProcess(
            args=args[0],
            returncode=1,
            stdout="ERROR: invalid manifest entry at META-INF/manifest.xml",
            stderr="",
        )

    monkeypatch.setattr(module.subprocess, "run", fake_run)

    result = module._run_reference_runner(
        module.RunnerConfig(name="opf", command_template=["opf-validator", "{file}"]),
        sample,
        timeout_seconds=10,
    )

    assert result["status"] == "ok"
    assert result["issue_count"] >= 1
    assert result["issues"]


def test_run_reference_runner_marks_docker_daemon_permission_unavailable(
    tmp_path: Path, monkeypatch
) -> None:
    module = _load_runner_module()
    sample = tmp_path / "sample.odt"
    sample.write_text("stub", encoding="utf-8")

    def fake_run(*args, **kwargs):  # type: ignore[no-untyped-def]
        return subprocess.CompletedProcess(
            args=args[0],
            returncode=126,
            stdout="",
            stderr=(
                "docker: permission denied while trying to connect to the Docker daemon socket.\n"
            ),
        )

    monkeypatch.setattr(module.subprocess, "run", fake_run)

    result = module._run_reference_runner(
        module.RunnerConfig(name="odf_toolkit", command_template=["docker", "run", "{file}"]),
        sample,
        timeout_seconds=10,
    )

    assert result["status"] == "unavailable"
    assert result["issue_count"] == 0
