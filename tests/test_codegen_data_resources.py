"""Tests for Open XML SDK data resource resolution."""

from __future__ import annotations

import json
from pathlib import Path

import openxml_audit.codegen.data_resources as data_resources
import openxml_audit.codegen.schema_loader as schema_loader
import openxml_audit.codegen.schematron_loader as schematron_loader
from tests.bridge_invariant_helpers import (
    collect_live_enum_type_names,
    collect_live_enum_type_owners,
)


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


def test_get_enum_values_falls_back_to_schema_enum_definitions(
    tmp_path: Path,
    monkeypatch,
) -> None:
    """Enum values should still resolve from schema JSON when enums.json is absent."""
    data_dir = tmp_path / "openxml"
    _write_json(
        data_dir / "schemas" / "presentation.json",
        {
            "TargetNamespace": "urn:test:presentation",
            "Types": [],
            "Enums": [
                {
                    "Type": "p:ST_TLTimeIndefinite",
                    "Name": "IndefiniteTimeDeclarationValues",
                    "Facets": [
                        {
                            "Value": "indefinite",
                        }
                    ],
                }
            ],
        },
    )

    monkeypatch.setattr(schema_loader, "get_openxml_data_dir", lambda: data_dir)
    monkeypatch.setattr(schema_loader, "_enum_values", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_type", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_name", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_candidates_by_name", None)

    assert schema_loader.get_enum_values("p:ST_TLTimeIndefinite") == ["indefinite"]
    assert schema_loader.get_enum_values(
        "DocumentFormat.OpenXml.Presentation.IndefiniteTimeDeclarationValues"
    ) == ["indefinite"]


def test_get_enum_values_uses_dotnet_namespace_to_resolve_ambiguous_enum_names(
    tmp_path: Path,
    monkeypatch,
) -> None:
    data_dir = tmp_path / "openxml"
    _write_json(
        data_dir / "schemas" / "ambiguous.json",
        {
            "TargetNamespace": "urn:test",
            "Types": [],
            "Enums": [
                {
                    "Type": "x:ST_Type",
                    "Name": "RuleValues",
                    "Facets": [
                        {
                            "Value": "none",
                        },
                        {
                            "Value": "all",
                        },
                    ],
                },
                {
                    "Type": "o:ST_RType",
                    "Name": "RuleValues",
                    "Facets": [
                        {
                            "Value": "arc",
                        },
                        {
                            "Value": "callout",
                        },
                    ],
                },
                {
                    "Type": "a:ST_ColorSchemeIndex",
                    "Name": "ColorSchemeIndexValues",
                    "Facets": [
                        {
                            "Value": "dk1",
                        },
                        {
                            "Value": "lt1",
                        },
                    ],
                },
                {
                    "Type": "w:ST_ColorSchemeIndex",
                    "Name": "ColorSchemeIndexValues",
                    "Facets": [
                        {
                            "Value": "dark1",
                        },
                        {
                            "Value": "light1",
                        },
                    ],
                },
            ],
        },
    )

    monkeypatch.setattr(schema_loader, "get_openxml_data_dir", lambda: data_dir)
    monkeypatch.setattr(schema_loader, "_enum_values", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_type", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_name", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_candidates_by_name", None)

    assert schema_loader.get_enum_values("DocumentFormat.OpenXml.Spreadsheet.RuleValues") == [
        "none",
        "all",
    ]
    assert schema_loader.get_enum_values("DocumentFormat.OpenXml.Vml.Office.RuleValues") == [
        "arc",
        "callout",
    ]
    assert schema_loader.get_enum_values(
        "DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues"
    ) == ["dk1", "lt1"]
    assert schema_loader.get_enum_values(
        "DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues"
    ) == ["dark1", "light1"]


def test_get_enum_values_uses_dotnet_namespace_to_resolve_three_way_short_name_collisions(
    tmp_path: Path,
    monkeypatch,
) -> None:
    data_dir = tmp_path / "openxml"
    _write_json(
        data_dir / "schemas" / "ambiguous.json",
        {
            "TargetNamespace": "urn:test",
            "Types": [],
            "Enums": [
                {
                    "Type": "w:ST_StyleType",
                    "Name": "StyleValues",
                    "Facets": [
                        {"Value": "paragraph"},
                        {"Value": "character"},
                        {"Value": "table"},
                        {"Value": "numbering"},
                    ],
                },
                {
                    "Type": "w14:ST_Style",
                    "Name": "StyleValues",
                    "Facets": [
                        {"Value": "normal"},
                        {"Value": "warning"},
                        {"Value": "error"},
                    ],
                },
                {
                    "Type": "m:ST_Style",
                    "Name": "StyleValues",
                    "Facets": [
                        {"Value": "p"},
                        {"Value": "b"},
                        {"Value": "i"},
                        {"Value": "bi"},
                    ],
                },
            ],
        },
    )

    monkeypatch.setattr(schema_loader, "get_openxml_data_dir", lambda: data_dir)
    monkeypatch.setattr(schema_loader, "_enum_values", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_type", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_name", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_candidates_by_name", None)

    assert schema_loader.get_enum_values("DocumentFormat.OpenXml.Wordprocessing.StyleValues") == [
        "paragraph",
        "character",
        "table",
        "numbering",
    ]
    assert schema_loader.get_enum_values("DocumentFormat.OpenXml.Office2010.Word.StyleValues") == [
        "normal",
        "warning",
        "error",
    ]
    assert schema_loader.get_enum_values("DocumentFormat.OpenXml.Math.StyleValues") == [
        "p",
        "b",
        "i",
        "bi",
    ]


def test_get_enum_values_resolves_diagram_and_customui_short_name_collisions(
    tmp_path: Path,
    monkeypatch,
) -> None:
    data_dir = tmp_path / "openxml"
    _write_json(
        data_dir / "schemas" / "ambiguous.json",
        {
            "TargetNamespace": "urn:test",
            "Types": [],
            "Enums": [
                {
                    "Type": "dgm:ST_Direction",
                    "Name": "DirectionValues",
                    "Facets": [
                        {"Value": "norm"},
                        {"Value": "rev"},
                    ],
                },
                {
                    "Type": "p:ST_Direction",
                    "Name": "DirectionValues",
                    "Facets": [
                        {"Value": "horz"},
                        {"Value": "vert"},
                    ],
                },
                {
                    "Type": "w:ST_Direction",
                    "Name": "DirectionValues",
                    "Facets": [
                        {"Value": "ltr"},
                        {"Value": "rtl"},
                    ],
                },
                {
                    "Type": "mso14:ST_Style",
                    "Name": "StyleValues",
                    "Facets": [
                        {"Value": "normal"},
                        {"Value": "warning"},
                        {"Value": "error"},
                    ],
                },
                {
                    "Type": "w:ST_StyleType",
                    "Name": "StyleValues",
                    "Facets": [
                        {"Value": "paragraph"},
                        {"Value": "character"},
                    ],
                },
                {
                    "Type": "m:ST_Style",
                    "Name": "StyleValues",
                    "Facets": [
                        {"Value": "p"},
                        {"Value": "b"},
                    ],
                },
            ],
        },
    )

    monkeypatch.setattr(schema_loader, "get_openxml_data_dir", lambda: data_dir)
    monkeypatch.setattr(schema_loader, "_enum_values", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_type", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_name", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_candidates_by_name", None)

    assert schema_loader.get_enum_values(
        "DocumentFormat.OpenXml.Drawing.Diagrams.DirectionValues"
    ) == ["norm", "rev"]
    assert schema_loader.get_enum_values(
        "DocumentFormat.OpenXml.Office2010.CustomUI.StyleValues"
    ) == ["normal", "warning", "error"]


def test_shipped_schema_enum_values_resolve_all_live_enum_types(monkeypatch) -> None:
    monkeypatch.setattr(schema_loader, "_enum_values", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_type", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_values_by_name", None)
    monkeypatch.setattr(schema_loader, "_schema_enum_candidates_by_name", None)

    owners = collect_live_enum_type_owners()
    unresolved = [
        (
            enum_type,
            owners.get(enum_type, ()),
        )
        for enum_type in collect_live_enum_type_names()
        if schema_loader.get_enum_values(enum_type) is None
    ]
    assert unresolved == []


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
