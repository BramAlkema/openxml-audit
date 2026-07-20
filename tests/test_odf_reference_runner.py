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


def test_run_reference_runner_runs_command_in_configured_working_dir(
    tmp_path: Path, monkeypatch
) -> None:
    # Regression guard: the OPF launcher resolves its jar via a repo-relative
    # path, so it must execute from the checkout root. Without cwd every sample
    # failed with "Unable to access jarfile" and was classified "unavailable".
    module = _load_runner_module()
    sample = tmp_path / "sample.odt"
    sample.write_text("stub", encoding="utf-8")
    checkout_root = tmp_path / "opf_checkout"
    checkout_root.mkdir()

    captured: dict[str, object] = {}

    def fake_run(*args, **kwargs):  # type: ignore[no-untyped-def]
        captured["cwd"] = kwargs.get("cwd")
        return subprocess.CompletedProcess(
            args=args[0], returncode=0, stdout="{}", stderr=""
        )

    monkeypatch.setattr(module.subprocess, "run", fake_run)

    result = module._run_reference_runner(
        module.RunnerConfig(
            name="opf",
            command_template=["odf-validator", "{file}"],
            working_dir=str(checkout_root),
        ),
        sample,
        timeout_seconds=10,
    )

    assert captured["cwd"] == str(checkout_root)
    assert result["status"] == "ok"


def test_run_reference_runner_defaults_to_no_working_dir(
    tmp_path: Path, monkeypatch
) -> None:
    module = _load_runner_module()
    sample = tmp_path / "sample.odt"
    sample.write_text("stub", encoding="utf-8")

    captured: dict[str, object] = {}

    def fake_run(*args, **kwargs):  # type: ignore[no-untyped-def]
        captured["cwd"] = kwargs.get("cwd", "MISSING")
        return subprocess.CompletedProcess(
            args=args[0], returncode=0, stdout="{}", stderr=""
        )

    monkeypatch.setattr(module.subprocess, "run", fake_run)

    module._run_reference_runner(
        module.RunnerConfig(name="odf_toolkit", command_template=["java", "{file}"]),
        sample,
        timeout_seconds=10,
    )

    # No working_dir configured -> cwd left unset (None), preserving prior behavior.
    assert captured["cwd"] is None


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


def test_reference_issue_keys_are_stable_across_staging_directories(
    tmp_path: Path,
) -> None:
    # The comparison key embeds the path the reference validator echoes back.
    # Locally that path lives under a per-run temporary directory, so without
    # canonicalisation every family looks new on every run and no baseline can
    # match. Two runs of the same sample from different staging roots must
    # produce identical keys.
    module = _load_runner_module()

    def keys_for(staging_root: Path) -> list[str]:
        sample_dir = staging_root / "invalid_mimetype"
        sample_dir.mkdir(parents=True)
        staged = sample_dir / "invalid_mimetype.odt"
        staged.write_bytes(b"")
        message = (
            f"{staged}/mimetype: Error: mimetype is not an ODFMediaTypes mimetype."
        )
        rows = module._reference_issue_rows("odf_toolkit", [message], staged)
        return [row["comparison_key"] for row in rows]

    first = keys_for(tmp_path / "run_a1b2c3")
    second = keys_for(tmp_path / "run_z9y8x7")

    assert first == second
    assert first, "expected at least one issue row"
    assert "/input/invalid_mimetype.odt" in first[0]
    assert "run_a1b2c3" not in first[0]


def test_reference_issue_rows_without_a_path_are_unchanged(tmp_path: Path) -> None:
    module = _load_runner_module()
    message = "Error: mimetype is not an ODFMediaTypes mimetype."
    rows = module._reference_issue_rows("odf_toolkit", [message], None)
    assert rows and "mimetype" in rows[0]["comparison_key"]
