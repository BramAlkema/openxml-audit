# ODF Security Policy

This document defines what `openxml-audit` security-core checks do and do not guarantee.

## Scope

Security-core is opt-in via the ODF validator:

```python
from openxml_audit.odf import OdfValidator

validator = OdfValidator(
    schema_validation=True,
    semantic_validation=True,
    security_validation=True,
)
```

Cryptographic verification is also optional and dependency-gated:

```python
validator = OdfValidator(
    security_validation=True,
    verify_cryptography=True,
)
```

If no verifier backend is available, a policy diagnostic is emitted.

## What Security-Core Validates

### Signature structure

- `META-INF/documentsignatures.xml` manifest entry media type (`ODFSEC001`)
- signature package root element namespace/name (`ODFSEC002`)
- presence of at least one `ds:Signature` (`ODFSEC003`)
- basic `ds:SignedInfo` + `ds:Reference` shape (`ODFSEC004`)

### Encryption structure

- root entry `'/'` cannot carry `manifest:encryption-data` (`ODFSEC101`)
- encrypted members must declare algorithm metadata (`ODFSEC102`)
- encrypted members must declare key-derivation metadata (`ODFSEC103`)
- checksum fields must be complete (`ODFSEC104`)

### Cryptographic verification policy

- requested verification without an available backend is reported (`ODFSEC900`)
- backend verification issues are reported (`ODFSEC901`)

## Non-Guarantees

Security-core does not guarantee:

- trust-chain validity
- certificate revocation status
- signature timestamp validity
- enterprise key-management policy compliance
- payload confidentiality strength assessment

These require environment-specific crypto policy, trust stores, and potentially external services.

## Dependency-Gated Hook Path

A custom verifier can be injected with `cryptographic_verifier`.

Built-in default verifier loading is best-effort and currently attempts optional `signxml`.
Install with:

```bash
pip install -e ".[odf-crypto]"
```

If unavailable, validation continues with structural checks and emits `ODFSEC900` when crypto verification is requested.
