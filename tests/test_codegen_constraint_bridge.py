"""Tests for SDK-to-runtime validator bridging."""

from __future__ import annotations

import json
from pathlib import Path

from lxml import etree

import openxml_audit.codegen.constraint_bridge as constraint_bridge_module
import openxml_audit.codegen.schema_loader as schema_loader_module
from openxml_audit.codegen.constraint_bridge import _build_type_validator
from openxml_audit.codegen.schema_loader import SdkAttribute
from openxml_audit.context import ValidationContext
from openxml_audit.errors import FileFormat


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload), encoding="utf-8")


def test_build_type_validator_uses_enum_validator_type_for_string_unions(
    monkeypatch,
) -> None:
    """Enum validators on StringValue attributes should use the validator type."""
    monkeypatch.setattr(
        constraint_bridge_module,
        "get_enum_values",
        lambda enum_type: ["indefinite"] if enum_type == "p:ST_TLTimeIndefinite" else None,
    )

    attr = SdkAttribute(
        qname=":tm",
        property_name="Time",
        type_name="StringValue",
        validators=[
            {
                "Name": "NumberValidator",
                "Type": "xsd:unsignedInt",
                "Version": "Office2007",
                "UnionId": 0,
            },
            {
                "Name": "EnumValidator",
                "Type": "p:ST_TLTimeIndefinite",
                "Version": "Office2007",
                "UnionId": 0,
            },
        ],
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    context = ValidationContext(file_format=FileFormat.OFFICE_2007)
    assert validator.validate("indefinite", context).is_valid
    assert validator.validate("12", context).is_valid
    assert not validator.validate("-1", context).is_valid


def test_build_type_validator_selects_versioned_snapshot_by_file_format() -> None:
    """Later SDK validator snapshots should not leak into earlier file formats."""
    attr = SdkAttribute(
        qname="w:w",
        property_name="Width",
        type_name="StringValue",
        validators=[
            {
                "Name": "NumberValidator",
                "Type": "w:ST_DecimalNumber",
                "Version": "Office2007",
            },
            {
                "Arguments": [
                    {
                        "Type": "String",
                        "Name": "Pattern",
                        "Value": r"-?[0-9]+(\.[0-9]+)?%",
                    }
                ],
                "Name": "StringValidator",
                "Version": "Office2010",
                "UnionId": 0,
            },
            {
                "Name": "NumberValidator",
                "Type": "w:ST_DecimalNumber",
                "Version": "Office2010",
                "UnionId": 0,
            },
        ],
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    office_2007 = ValidationContext(file_format=FileFormat.OFFICE_2007)
    office_2010 = ValidationContext(file_format=FileFormat.OFFICE_2010)

    assert validator.validate("25", office_2007).is_valid
    assert not validator.validate("25%", office_2007).is_valid
    assert validator.validate("25%", office_2010).is_valid


def test_build_type_validator_preserves_decimal_number_constraints() -> None:
    """Decimal-valued NumberValidator rules should not be coerced to integers."""
    attr = SdkAttribute(
        qname="emma:confidence",
        property_name="Confidence",
        type_name="DecimalValue",
        validators=[
            {
                "Arguments": [
                    {
                        "Type": "Long",
                        "Name": "MinInclusive",
                        "Value": "0",
                    },
                    {
                        "Type": "Long",
                        "Name": "MaxInclusive",
                        "Value": "1",
                    },
                ],
                "Name": "NumberValidator",
            }
        ],
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    context = ValidationContext()
    assert validator.validate("0.5", context).is_valid
    assert validator.validate("1", context).is_valid
    assert not validator.validate("2", context).is_valid


def test_build_type_validator_uses_schema_enum_fallback_for_ambiguous_dotnet_enum(
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
            ],
        },
    )

    monkeypatch.setattr(schema_loader_module, "get_openxml_data_dir", lambda: data_dir)
    monkeypatch.setattr(schema_loader_module, "_enum_values", None)
    monkeypatch.setattr(schema_loader_module, "_schema_enum_values_by_type", None)
    monkeypatch.setattr(schema_loader_module, "_schema_enum_values_by_name", None)
    monkeypatch.setattr(schema_loader_module, "_schema_enum_candidates_by_name", None)

    attr = SdkAttribute(
        qname=":type",
        property_name="Type",
        type_name="EnumValue<DocumentFormat.OpenXml.Spreadsheet.RuleValues>",
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    context = ValidationContext(file_format=FileFormat.OFFICE_2019)
    assert validator.validate("none", context).is_valid
    assert not validator.validate("arc", context).is_valid
    assert not validator.validate("definitely-not-a-rule", context).is_valid


def test_build_type_validator_preserves_enum_type_with_string_metadata(monkeypatch) -> None:
    monkeypatch.setattr(
        constraint_bridge_module,
        "get_enum_values",
        lambda enum_type: (
            ["auto", "gray"]
            if enum_type == "DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues"
            else None
        ),
    )

    attr = SdkAttribute(
        qname=":bwMode",
        property_name="BlackWhiteMode",
        type_name="EnumValue<DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues>",
        validators=[
            {
                "Name": "StringValidator",
                "Arguments": [
                    {
                        "Name": "IsToken",
                        "Value": "True",
                    }
                ],
            }
        ],
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    context = ValidationContext()
    assert validator.validate("auto", context).is_valid
    assert not validator.validate("bogus", context).is_valid


def test_build_type_validator_preserves_list_enum_type_from_sdk_type_name(monkeypatch) -> None:
    monkeypatch.setattr(
        constraint_bridge_module,
        "get_enum_values",
        lambda enum_type: (
            ["acoustic", "tactile"]
            if enum_type == "DocumentFormat.OpenXml.EMMA.MediumValues"
            else None
        ),
    )

    attr = SdkAttribute(
        qname="emma:medium",
        property_name="Medium",
        type_name="ListValue<EnumValue<DocumentFormat.OpenXml.EMMA.MediumValues>>",
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    context = ValidationContext()
    assert validator.validate("acoustic tactile", context).is_valid
    assert not validator.validate("acoustic bogus", context).is_valid


def test_build_type_validator_preserves_number_list_validators() -> None:
    attr = SdkAttribute(
        qname=":ascender",
        property_name="Ascender",
        type_name="StringValue",
        validators=[
            {
                "Name": "NumberValidator",
                "Type": "msink:ST_Point",
                "UnionId": 0,
                "IsList": True,
            },
            {
                "Name": "NumberValidator",
                "Type": "xsd:int",
                "UnionId": 0,
            },
        ],
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    context = ValidationContext()
    assert validator.validate("5", context).is_valid
    assert validator.validate("1 2", context).is_valid
    assert not validator.validate("1 two", context).is_valid


def test_build_type_validator_honors_qname_string_flags() -> None:
    attr = SdkAttribute(
        qname=":idQ",
        property_name="IdQ",
        type_name="StringValue",
        validators=[
            {
                "Name": "StringValidator",
                "Arguments": [
                    {
                        "Name": "IsQName",
                        "Value": "True",
                    }
                ],
            }
        ],
    )

    validator = _build_type_validator(attr)
    assert validator is not None

    element = etree.fromstring(b'<mso:control xmlns:mso="urn:test"/>')
    context = ValidationContext()
    context.push_element(element)
    try:
        assert validator.validate("mso:button", context).is_valid
        assert not validator.validate("missing:button", context).is_valid
        assert not validator.validate("not a qname", context).is_valid
    finally:
        context.pop_element()
