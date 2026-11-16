from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType

SCRIPT_PATH = (
    Path(__file__).resolve().parents[1] / "scripts" / "odf" / "bootstrap_reference_validators.py"
)


def _load_bootstrap_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location("odf_bootstrap", SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load bootstrap script module from {SCRIPT_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_select_odf_toolkit_jar_prefers_fat_jar(tmp_path: Path) -> None:
    module = _load_bootstrap_module()
    target_dir = tmp_path / "validator" / "target"
    target_dir.mkdir(parents=True)

    plain_jar = target_dir / "odfvalidator-0.13.0.jar"
    fat_jar = target_dir / "odfvalidator-0.13.0-jar-with-dependencies.jar"
    plain_jar.write_text("plain", encoding="utf-8")
    fat_jar.write_text("fat", encoding="utf-8")

    selected = module._select_odf_toolkit_jar(tmp_path)
    assert selected == fat_jar


def test_select_opf_launcher_prefers_script(tmp_path: Path) -> None:
    module = _load_bootstrap_module()
    launcher = tmp_path / "odf-validator"
    launcher.write_text("#!/bin/sh\n", encoding="utf-8")
    target_dir = tmp_path / "target"
    target_dir.mkdir()
    (target_dir / "opf-validator.jar").write_text("jar", encoding="utf-8")

    kind, path = module._select_opf_launcher(tmp_path)
    assert kind == "script"
    assert path == launcher


def test_select_opf_launcher_falls_back_to_jar(tmp_path: Path) -> None:
    module = _load_bootstrap_module()
    target_dir = tmp_path / "target"
    target_dir.mkdir()
    jar = target_dir / "opf-validator.jar"
    jar.write_text("jar", encoding="utf-8")

    kind, path = module._select_opf_launcher(tmp_path)
    assert kind == "jar"
    assert path == jar


def test_resolve_runtime_mode_auto_uses_docker_when_java_is_unavailable(monkeypatch) -> None:
    module = _load_bootstrap_module()

    monkeypatch.setattr(module, "_java_runtime_ready", lambda: False)
    monkeypatch.setattr(
        module.shutil,
        "which",
        lambda name: "/usr/bin/docker" if name == "docker" else None,
    )

    assert module._resolve_runtime_mode("auto") == "docker"


def test_build_odf_toolkit_command_template_docker_uses_file_placeholders(
    tmp_path: Path,
) -> None:
    module = _load_bootstrap_module()
    target = tmp_path / "odf_toolkit"
    jar = target / "validator" / "target" / "odfvalidator-fat.jar"
    jar.parent.mkdir(parents=True)
    jar.write_text("jar", encoding="utf-8")

    template = module._build_odf_toolkit_command_template(
        target=target,
        jar_path=jar,
        runtime_mode="docker",
        runtime_image="eclipse-temurin:17-jre",
    )

    assert "docker run --rm" in template
    assert "{file_dir}:/input:ro" in template
    assert "/input/{file_name}" in template
