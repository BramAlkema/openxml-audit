"""Verify ValidationError.source_class is set correctly per finding origin.

Spec 018 (0.6.2): each finder tags its emissions with the right SourceClass
so downstream parity tooling can filter (e.g. exclude WORD_APP_COMPAT findings
from the advisory SDK comparison; include them in the future self-parity gate).
"""

from __future__ import annotations

from openxml_audit.context import ValidationContext
from openxml_audit.errors import (
    SourceClass,
    ValidationError,
    ValidationErrorType,
)


def test_validation_error_default_is_sdk_proxy() -> None:
    """ValidationError's default source_class is SDK_PROXY (covers schema/semantic
    findings that mirror the .NET SDK)."""
    err = ValidationError(
        error_type=ValidationErrorType.SCHEMA,
        description="anything",
    )
    assert err.source_class is SourceClass.SDK_PROXY


def test_context_add_error_passes_source_class_through() -> None:
    """ValidationContext.add_error accepts source_class and propagates it."""
    context = ValidationContext()
    context.add_error(
        error_type=ValidationErrorType.SEMANTIC,
        description="word repair-dialog finding",
        source_class=SourceClass.WORD_APP_COMPAT,
    )
    assert len(context.errors) == 1
    assert context.errors[0].source_class is SourceClass.WORD_APP_COMPAT


def test_context_add_schema_error_defaults_to_sdk_proxy() -> None:
    """add_schema_error keeps SDK_PROXY as the default."""
    context = ValidationContext()
    context.add_schema_error("schema thing")
    assert context.errors[0].source_class is SourceClass.SDK_PROXY


def test_source_class_enum_values() -> None:
    """The set of source classes is the documented six."""
    assert {sc.value for sc in SourceClass} == {
        "sdk_proxy",
        "word_app_compat",
        "excel_app_compat",
        "powerpoint_app_compat",
        "odf_native",
        "package_integrity",
    }


def test_source_class_exported_from_package() -> None:
    """SourceClass is importable from the top-level openxml_audit package."""
    import openxml_audit

    assert openxml_audit.SourceClass is SourceClass
