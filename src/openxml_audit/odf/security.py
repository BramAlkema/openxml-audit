"""ODF security-core validation (signature and encryption structure)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import (
    SIGNATURE_PKG_NS,
    XMLDSIG_NS,
    normalize_manifest_path,
    normalize_part_uri,
)
from openxml_audit.odf.package import OdfManifestEntry, OdfPackage


@dataclass(frozen=True)
class OdfSecurityRule:
    """Stable security-core rule metadata."""

    id: str
    family: str
    description: str


RULES: tuple[OdfSecurityRule, ...] = (
    OdfSecurityRule(
        id="ODFSEC001",
        family="signature",
        description="Signature manifest entry should declare text/xml media type.",
    ),
    OdfSecurityRule(
        id="ODFSEC002",
        family="signature",
        description=(
            "Signature package root must be document-signatures "
            "in ODF signature namespace."
        ),
    ),
    OdfSecurityRule(
        id="ODFSEC003",
        family="signature",
        description="documentsignatures.xml must contain at least one ds:Signature.",
    ),
    OdfSecurityRule(
        id="ODFSEC004",
        family="signature",
        description="Each ds:Signature should include ds:SignedInfo and at least one ds:Reference.",
    ),
    OdfSecurityRule(
        id="ODFSEC101",
        family="encryption",
        description="Encrypted file-entry must not be the manifest root entry '/'.",
    ),
    OdfSecurityRule(
        id="ODFSEC102",
        family="encryption",
        description="Encrypted file-entry must include manifest:algorithm metadata.",
    ),
    OdfSecurityRule(
        id="ODFSEC103",
        family="encryption",
        description="Encrypted file-entry must include manifest:key-derivation metadata.",
    ),
    OdfSecurityRule(
        id="ODFSEC104",
        family="encryption",
        description="manifest:checksum and manifest:checksum-type should be provided together.",
    ),
    OdfSecurityRule(
        id="ODFSEC900",
        family="verification",
        description="Cryptographic verification requested but no verifier is available.",
    ),
    OdfSecurityRule(
        id="ODFSEC901",
        family="verification",
        description="Cryptographic verifier reported an issue.",
    ),
)


def get_odf_security_rules() -> tuple[OdfSecurityRule, ...]:
    """Return security-core rule metadata with stable identifiers."""
    return RULES


class OdfCryptographicVerifier(Protocol):
    """Optional hook for cryptographic verification backends."""

    def verify_signatures(
        self,
        package: OdfPackage,
        signatures_root: etree._Element,
    ) -> list[str]:
        """Verify signatures and return issue descriptions."""

    def verify_encryption(
        self,
        package: OdfPackage,
        encrypted_entries: list[OdfManifestEntry],
    ) -> list[str]:
        """Validate encryption metadata and return issue descriptions."""


class _SignxmlCryptographicVerifier:
    """Best-effort SignXML-backed verifier used only when dependency is available."""

    def __init__(self, xml_verifier_cls: type, invalid_signature_exc: type[Exception]) -> None:
        self._xml_verifier_cls = xml_verifier_cls
        self._invalid_signature_exc = invalid_signature_exc

    def verify_signatures(
        self,
        package: OdfPackage,
        signatures_root: etree._Element,
    ) -> list[str]:
        del package
        issues: list[str] = []
        signatures = signatures_root.xpath(
            ".//ds:Signature",
            namespaces={"ds": XMLDSIG_NS},
        )
        for signature in signatures:
            try:
                verifier = self._xml_verifier_cls()
                verifier.verify(signature)
            except self._invalid_signature_exc as exc:
                issues.append(f"Signature verification failed: {exc}")
            except Exception as exc:  # pragma: no cover - defensive fallback
                issues.append(f"Signature verification backend error: {exc}")
        return issues

    def verify_encryption(
        self,
        package: OdfPackage,
        encrypted_entries: list[OdfManifestEntry],
    ) -> list[str]:
        del package
        del encrypted_entries
        return []


def load_default_cryptographic_verifier() -> tuple[OdfCryptographicVerifier | None, str]:
    """Try to construct a default verifier from optional dependencies."""
    try:
        from signxml import InvalidSignature, XMLVerifier
    except Exception:
        return None, "signxml dependency is not installed"
    return _SignxmlCryptographicVerifier(XMLVerifier, InvalidSignature), ""


class OdfSecurityValidator:
    """Security-core checks for ODF signatures and encryption declarations."""

    SIGNATURE_PATH = "META-INF/documentsignatures.xml"

    def __init__(
        self,
        *,
        verify_cryptography: bool = False,
        cryptographic_verifier: OdfCryptographicVerifier | None = None,
    ) -> None:
        self._verify_cryptography = verify_cryptography
        self._cryptographic_verifier: OdfCryptographicVerifier | None = cryptographic_verifier
        self._default_verifier_reason = ""
        if self._verify_cryptography and self._cryptographic_verifier is None:
            verifier, reason = load_default_cryptographic_verifier()
            self._cryptographic_verifier = verifier
            self._default_verifier_reason = reason

    @staticmethod
    def _normalize_part_uri(part_path: str) -> str:
        return normalize_part_uri(part_path)

    @staticmethod
    def _normalize_manifest_path(path: str) -> str:
        return normalize_manifest_path(path)

    @staticmethod
    def _error(
        *,
        rule_id: str,
        description: str,
        part_uri: str,
        error_type: ValidationErrorType = ValidationErrorType.SEMANTIC,
        severity: ValidationSeverity = ValidationSeverity.ERROR,
    ) -> ValidationError:
        return ValidationError(
            error_type=error_type,
            description=description,
            part_uri=part_uri,
            severity=severity,
            id=rule_id,
        )

    def validate(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        encrypted_entries = self._encrypted_entries(package)
        errors.extend(self._validate_signature_structure(package, parsed_parts))
        errors.extend(self._validate_encryption_structure(encrypted_entries))
        errors.extend(
            self._run_cryptographic_verification(package, parsed_parts, encrypted_entries)
        )
        return errors

    def _manifest_entry_by_path(self, package: OdfPackage, path: str) -> OdfManifestEntry | None:
        for entry in package.manifest:
            if self._normalize_manifest_path(entry.full_path) == path:
                return entry
        return None

    def _validate_signature_structure(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        entry = self._manifest_entry_by_path(package, self.SIGNATURE_PATH)
        if entry is None:
            return errors

        media_type = entry.media_type.strip().lower()
        if media_type != "text/xml":
            errors.append(
                self._error(
                    rule_id="ODFSEC001",
                    description=(
                        "Signature manifest entry should use media-type 'text/xml' "
                        f"(found '{entry.media_type}')"
                    ),
                    part_uri="/META-INF/manifest.xml",
                )
            )

        signatures_root = parsed_parts.get(self.SIGNATURE_PATH)
        if signatures_root is None:
            return errors

        qname = etree.QName(signatures_root)
        if (
            qname.namespace != SIGNATURE_PKG_NS
            or qname.localname != "document-signatures"
        ):
            errors.append(
                self._error(
                    rule_id="ODFSEC002",
                    description=(
                        "documentsignatures.xml root must be "
                        "document-signatures in ODF signature namespace"
                    ),
                    part_uri=self._normalize_part_uri(self.SIGNATURE_PATH),
                    error_type=ValidationErrorType.SCHEMA,
                )
            )
            return errors

        signatures = signatures_root.xpath(
            ".//ds:Signature",
            namespaces={"ds": XMLDSIG_NS},
        )
        if not signatures:
            errors.append(
                self._error(
                    rule_id="ODFSEC003",
                    description="documentsignatures.xml contains no ds:Signature entries",
                    part_uri=self._normalize_part_uri(self.SIGNATURE_PATH),
                )
            )
            return errors

        for signature in signatures:
            signed_info = signature.find(f"{{{XMLDSIG_NS}}}SignedInfo")
            references = signature.findall(
                f"{{{XMLDSIG_NS}}}SignedInfo/{{{XMLDSIG_NS}}}Reference"
            )
            if signed_info is None or not references:
                errors.append(
                    self._error(
                        rule_id="ODFSEC004",
                        description=(
                            "Each ds:Signature should contain ds:SignedInfo "
                            "with at least one ds:Reference"
                        ),
                        part_uri=self._normalize_part_uri(self.SIGNATURE_PATH),
                    )
                )
        return errors

    def _encrypted_entries(self, package: OdfPackage) -> list[OdfManifestEntry]:
        return [entry for entry in package.manifest if entry.has_encryption_data]

    def _validate_encryption_structure(
        self, encrypted_entries: list[OdfManifestEntry]
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for entry in encrypted_entries:
            path = self._normalize_manifest_path(entry.full_path)
            part_uri = "/META-INF/manifest.xml"

            if path == "/":
                errors.append(
                    self._error(
                        rule_id="ODFSEC101",
                        description="Root manifest entry '/' must not carry encryption-data",
                        part_uri=part_uri,
                    )
                )
            if not entry.encryption_algorithm_name:
                errors.append(
                    self._error(
                        rule_id="ODFSEC102",
                        description=(
                            f"Encrypted manifest entry '{path}' is missing "
                            "manifest:algorithm-name"
                        ),
                        part_uri=part_uri,
                    )
                )
            if not entry.encryption_key_derivation_name:
                errors.append(
                    self._error(
                        rule_id="ODFSEC103",
                        description=(
                            f"Encrypted manifest entry '{path}' is missing "
                            "manifest:key-derivation-name"
                        ),
                        part_uri=part_uri,
                    )
                )

            has_checksum_type = bool(entry.encryption_checksum_type)
            has_checksum = bool(entry.encryption_checksum)
            if has_checksum_type != has_checksum:
                errors.append(
                    self._error(
                        rule_id="ODFSEC104",
                        description=(
                            f"Encrypted manifest entry '{path}' should provide both "
                            "manifest:checksum-type and manifest:checksum together"
                        ),
                        part_uri=part_uri,
                    )
                )
        return errors

    def _run_cryptographic_verification(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
        encrypted_entries: list[OdfManifestEntry],
    ) -> list[ValidationError]:
        if not self._verify_cryptography:
            return []

        verifier = self._cryptographic_verifier
        if verifier is None:
            reason = self._default_verifier_reason or "no cryptographic verifier configured"
            return [
                self._error(
                    rule_id="ODFSEC900",
                    description=(
                        "Cryptographic verification was requested but is unavailable "
                        f"({reason}). Provide a custom cryptographic_verifier."
                    ),
                    part_uri="/",
                    severity=ValidationSeverity.WARNING,
                )
            ]

        errors: list[ValidationError] = []
        signatures_root = parsed_parts.get(self.SIGNATURE_PATH)
        if signatures_root is not None:
            try:
                issues = verifier.verify_signatures(package, signatures_root)
            except Exception as exc:  # pragma: no cover - defensive fallback
                issues = [f"Signature verifier hook failed: {exc}"]
            for issue in issues:
                errors.append(
                    self._error(
                        rule_id="ODFSEC901",
                        description=issue,
                        part_uri=self._normalize_part_uri(self.SIGNATURE_PATH),
                    )
                )

        if encrypted_entries:
            try:
                issues = verifier.verify_encryption(package, encrypted_entries)
            except Exception as exc:  # pragma: no cover - defensive fallback
                issues = [f"Encryption verifier hook failed: {exc}"]
            for issue in issues:
                errors.append(
                    self._error(
                        rule_id="ODFSEC901",
                        description=issue,
                        part_uri="/META-INF/manifest.xml",
                    )
                )

        return errors
