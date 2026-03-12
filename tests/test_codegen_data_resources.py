"""Tests for Open XML SDK data resource resolution."""

from __future__ import annotations

import json
from pathlib import Path

import openxml_audit.codegen.data_resources as data_resources
import openxml_audit.codegen.schema_loader as schema_loader
import openxml_audit.codegen.schematron_loader as schematron_loader


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload), encoding="utf-8")


def test_get_openxml_data_dir_prefers_packaged_resources(
    tmp_path: Path,
    monkeypatch,
) -> None:
    """Package resources should win when the wheel bundles the data."""
    package_root = tmp_path / "openxml_audit"
    packaged_data_dir = package_root / "data" / "openxml"
    packaged_data_dir.mkdir(parents=True)

    fallback_data_dir = tmp_path / "project-data" / "openxml"
    fallback_data_dir.mkdir(parents=True)

    monkeypatch.setattr(data_resources, "files", lambda package: package_root)
    monkeypatch.setattr(data_resources, "PROJECT_DATA_DIR", fallback_data_dir)

    assert data_resources.get_openxml_data_dir() == packaged_data_dir


def test_get_openxml_data_dir_falls_back_to_project_data_dir(
    tmp_path: Path,
    monkeypatch,
) -> None:
    """Source checkouts should continue to use the repo data directory."""
    package_root = tmp_path / "openxml_audit"
    package_root.mkdir()

    fallback_data_dir = tmp_path / "project-data" / "openxml"
    fallback_data_dir.mkdir(parents=True)

    monkeypatch.setattr(data_resources, "files", lambda package: package_root)
    monkeypatch.setattr(data_resources, "PROJECT_DATA_DIR", fallback_data_dir)

    assert data_resources.get_openxml_data_dir() == fallback_data_dir


def test_schema_registry_loads_from_resolved_data_dir(
    tmp_path: Path,
    monkeypatch,
) -> None:
    """Schema registry should use the resolved runtime data directory."""
    data_dir = tmp_path / "openxml"
    _write_json(
        data_dir / "namespaces.json",
        [{"Prefix": "p", "Uri": "urn:test:presentation"}],
    )
    _write_json(
        data_dir / "schemas" / "presentation.json",
        {
            "TargetNamespace": "urn:test:presentation",
            "Types": [
                {
                    "Name": "p:CT_Test/p:test",
                    "ClassName": "TestElement",
                }
            ],
        },
    )

    monkeypatch.setattr(schema_loader, "get_openxml_data_dir", lambda: data_dir)

    registry = schema_loader.SchemaRegistry()
    registry.load()

    assert registry.get_namespace("p") == "urn:test:presentation"
    assert registry.count_types() == 1
    assert registry.count_elements() == 1
    assert registry.get_element_type("urn:test:presentation", "test") is not None


def test_schematron_registry_loads_from_resolved_data_dir(
    tmp_path: Path,
    monkeypatch,
) -> None:
    """Schematron registry should use the resolved runtime data directory."""
    data_dir = tmp_path / "openxml"
    _write_json(
        data_dir / "schematrons.json",
        [
            {
                "Context": "p:test",
                "Test": "@id = 'demo'",
                "App": "All",
            }
        ],
    )

    monkeypatch.setattr(schematron_loader, "get_openxml_data_dir", lambda: data_dir)

    registry = schematron_loader.SchematronRegistry()
    registry.load()

    assert registry.count_rules() == 1
    rules = registry.get_rules_for_context("p:test")
    assert len(rules) == 1
    assert rules[0].expected_value == "demo"
