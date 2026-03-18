from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _load_script_module(name: str, relative_path: str) -> ModuleType:
    script_path = PROJECT_ROOT / relative_path
    spec = importlib.util.spec_from_file_location(name, script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load script module from {script_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_sync_openxml_data_persists_ref_and_commit(tmp_path: Path, monkeypatch) -> None:
    module = _load_script_module("sync_openxml_data_test", "scripts/sync_openxml_data.py")
    monkeypatch.setattr(module, "DATA_DIR", tmp_path / "data" / "openxml")

    module.save_version("abcdef1234567890", sdk_ref="v9.9.9")

    assert module.get_current_ref() == "v9.9.9"
    assert module.get_current_version() == "abcdef1234567890"


def test_check_sdk_update_tracks_sync_and_parity_gate_refs() -> None:
    module = _load_script_module("check_sdk_update_test", "scripts/check_sdk_update.py")

    assert ".github/workflows/parity-gate.yml" in module.PINNED_FILES
    assert "scripts/sync_openxml_data.py" in module.PINNED_FILES
