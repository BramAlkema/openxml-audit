"""Verify ValidationError.source_class is set correctly per finding origin.

Spec 018 (0.6.2): each finder tags its emissions with the right SourceClass
so downstream parity tooling can filter (e.g. exclude WORD_APP_COMPAT findings
from the advisory SDK comparison; include them in the future self-parity gate).
"""

from __future__ import annotations

from openxml_audit.context import ElementContext, ValidationContext
from openxml_audit.errors import (
    SourceClass,
    ValidationError,
    ValidationErrorType,
    ValidationSeverity,
)
from openxml_audit.namespaces import WORDPROCESSINGML
from openxml_audit.word.compat import WordCompatValidator
from lxml import etree


def _build_trpr_with_reordered_children() -> etree._Element:
    """trPr where cantSplit appears after tblHeader (issue #3 repro)."""
    nsmap = {"w": WORDPROCESSINGML}
    body = etree.Element(f"{{{WORDPROCESSINGML}}}body", nsmap=nsmap)
    tbl = etree.SubElement(body, f"{{{WORDPROCESSINGML}}}tbl")
    tr = etree.SubElement(tbl, f"{{{WORDPROCESSINGML}}}tr")
    trpr = etree.SubElement(tr, f"{{{WORDPROCESSINGML}}}trPr")
    etree.SubElement(trpr, f"{{{WORDPROCESSINGML}}}tblHeader")
    etree.SubElement(trpr, f"{{{WORDPROCESSINGML}}}cantSplit")
    return etree.ElementTree(body).getroot()


class _FakePart:
    """Minimal stand-in for DocumentPart sufficient for compat.validate()."""

    def __init__(self, xml: etree._Element) -> None:
        self.xml = xml


def test_word_compat_emits_word_app_compat() -> None:
    """compat.WordCompatValidator tags ordering findings as WORD_APP_COMPAT."""
    body = _build_trpr_with_reordered_children()
    context = ValidationContext()
    validator = WordCompatValidator()

    errors = validator.validate(_FakePart(body), context)

    assert len(errors) >= 1, "expected at least one ordering finding"
    for err in errors:
        assert err.source_class is SourceClass.WORD_APP_COMPAT, (
            f"Word compat finding should be WORD_APP_COMPAT, got {err.source_class}"
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
