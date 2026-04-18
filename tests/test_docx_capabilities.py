"""DOCX capability registry is empty at kickoff; structural tests only."""

from __future__ import annotations

from openxml_audit.docx.capabilities import (
    check_capability,
    get_capability_finding,
    list_capability_findings,
)


def test_list_capability_findings_is_empty() -> None:
    assert list_capability_findings() == []


def test_get_capability_finding_returns_none_for_unknown() -> None:
    assert get_capability_finding("docx.field.toc") is None


def test_check_capability_reports_unknown() -> None:
    payload = check_capability("docx.field.toc")
    assert payload["known"] is False
