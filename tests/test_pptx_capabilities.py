from __future__ import annotations

from openxml_audit.evidence import EvidenceTier
from openxml_audit.pptx.capabilities import (
    check_capability,
    get_capability_finding,
    list_capability_findings,
)


def test_get_capability_finding_resolves_alias() -> None:
    finding = get_capability_finding("fade")

    assert finding is not None
    assert finding.key == "pptx.anim.effect.entr.fade"


def test_check_capability_reports_minimum_tier_match() -> None:
    payload = check_capability(
        "pptx.anim.effect.entr.fade",
        minimum_tier=EvidenceTier.LOADABLE,
    )

    assert payload["known"] is True
    assert payload["meets_minimum_tier"] is True


def test_check_capability_reports_unknown_key() -> None:
    payload = check_capability("pptx.unknown.feature")

    assert payload["known"] is False


def test_list_capability_findings_filters_by_prefix() -> None:
    findings = list_capability_findings("pptx.timing.")

    assert findings
    assert all(finding.key.startswith("pptx.timing.") for finding in findings)
