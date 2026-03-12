"""Tests for ODF security-core validation."""

from __future__ import annotations

from pathlib import Path

import openxml_audit.odf.security as security_module
from openxml_audit.odf import OdfValidator, get_odf_security_rules
from openxml_audit.odf.package import OdfManifestEntry, OdfPackage


def _security_validator(
    *,
    verify_cryptography: bool = False,
    cryptographic_verifier: security_module.OdfCryptographicVerifier | None = None,
) -> OdfValidator:
    return OdfValidator(
        schema_validation=False,
        semantic_validation=False,
        security_validation=True,
        verify_cryptography=verify_cryptography,
        cryptographic_verifier=cryptographic_verifier,
    )


def test_security_rule_registry_exposes_stable_unique_ids() -> None:
    rules = get_odf_security_rules()
    ids = [rule.id for rule in rules]
    assert ids
    assert len(ids) == len(set(ids))
    assert all(rule.id.startswith("ODFSEC") for rule in rules)


def test_foundation_mode_keeps_signed_stub_valid(minimal_odt_signed_stub: Path) -> None:
    result = OdfValidator(
        schema_validation=False,
        semantic_validation=False,
        security_validation=False,
    ).validate(minimal_odt_signed_stub)
    assert result.is_valid


def test_security_core_accepts_signed_stub(minimal_odt_signed_stub: Path) -> None:
    result = _security_validator().validate(minimal_odt_signed_stub)
    assert result.is_valid


def test_security_core_accepts_structural_signature_fixture(
    minimal_odt_signed_structural: Path,
) -> None:
    result = _security_validator().validate(minimal_odt_signed_structural)
    assert result.is_valid


def test_signature_bad_root_reports_rule_id(odf_signature_bad_root: Path) -> None:
    result = _security_validator().validate(odf_signature_bad_root)
    assert not result.is_valid
    assert any(error.id == "ODFSEC002" for error in result.errors)


def test_signature_bad_media_type_reports_rule_id(odf_signature_bad_media_type: Path) -> None:
    result = _security_validator().validate(odf_signature_bad_media_type)
    assert not result.is_valid
    assert any(error.id == "ODFSEC001" for error in result.errors)


def test_signature_missing_signedinfo_reports_rule_id(
    odf_signature_missing_signedinfo: Path,
) -> None:
    result = _security_validator().validate(odf_signature_missing_signedinfo)
    assert not result.is_valid
    assert any(error.id == "ODFSEC004" for error in result.errors)


def test_security_core_accepts_structural_encryption_fixture(
    minimal_odt_encrypted_structural: Path,
) -> None:
    result = _security_validator().validate(minimal_odt_encrypted_structural)
    assert result.is_valid


def test_security_core_accepts_encrypted_stub(
    minimal_odt_encrypted_stub: Path,
) -> None:
    result = _security_validator().validate(minimal_odt_encrypted_stub)
    assert result.is_valid


def test_encrypted_missing_key_derivation_reports_rule_id(
    odf_encrypted_missing_key_derivation: Path,
) -> None:
    result = _security_validator().validate(odf_encrypted_missing_key_derivation)
    assert not result.is_valid
    assert any(error.id == "ODFSEC103" for error in result.errors)


def test_encrypted_root_entry_encrypted_reports_rule_id(
    odf_encrypted_root_entry_encrypted: Path,
) -> None:
    result = _security_validator().validate(odf_encrypted_root_entry_encrypted)
    assert not result.is_valid
    assert any(error.id == "ODFSEC101" for error in result.errors)


def test_encrypted_checksum_partial_reports_rule_id(
    odf_encrypted_checksum_partial: Path,
) -> None:
    result = _security_validator().validate(odf_encrypted_checksum_partial)
    assert not result.is_valid
    assert any(error.id == "ODFSEC104" for error in result.errors)


def test_verify_cryptography_without_dependency_reports_policy_diagnostic(
    minimal_odt_signed_structural: Path,
    monkeypatch,
) -> None:
    monkeypatch.setattr(
        security_module,
        "load_default_cryptographic_verifier",
        lambda: (None, "mock dependency missing"),
    )
    validator = OdfValidator(
        schema_validation=False,
        semantic_validation=False,
        security_validation=True,
        verify_cryptography=True,
    )
    result = validator.validate(minimal_odt_signed_structural)
    assert any(error.id == "ODFSEC900" for error in result.errors)


class _DummyVerifier:
    def verify_signatures(self, package: OdfPackage, signatures_root) -> list[str]:
        del package
        del signatures_root
        return ["dummy signature verification issue"]

    def verify_encryption(
        self,
        package: OdfPackage,
        encrypted_entries: list[OdfManifestEntry],
    ) -> list[str]:
        del package
        if not encrypted_entries:
            return []
        return ["dummy encryption verification issue"]


def test_custom_cryptographic_verifier_issues_reported(
    minimal_odt_signed_structural: Path,
    minimal_odt_encrypted_structural: Path,
) -> None:
    verifier = _DummyVerifier()

    signed_result = _security_validator(
        verify_cryptography=True,
        cryptographic_verifier=verifier,
    ).validate(minimal_odt_signed_structural)
    assert any(error.id == "ODFSEC901" for error in signed_result.errors)

    encrypted_result = _security_validator(
        verify_cryptography=True,
        cryptographic_verifier=verifier,
    ).validate(minimal_odt_encrypted_structural)
    assert any(error.id == "ODFSEC901" for error in encrypted_result.errors)
