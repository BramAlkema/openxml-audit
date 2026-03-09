# Spec: ODF OASIS Conformance Gap Closure

## Status

Proposed (March 9, 2026)

## Problem

`openxml-audit` now provides strong ODF foundation validation (package integrity, XML parse sweep, selected
document semantics, optional Relax NG hook, and reference-comparison tooling), but it is still short of full
OASIS-level conformance checks.

Current gaps include:

1. Relax NG coverage is selective, not full-package.
2. Semantic rule coverage is narrow compared to ODF conformance requirements.
3. Signature and encryption handling is policy-level only, not full cryptographic validation.
4. Reference validator parity is tooling-capable but not enforced with fully wired external runs.

## Why This Matters

- "ODF support" can be interpreted as standards-level conformance.
- Foundation checks catch corruption, but not the full range of interoperability failures.
- A staged conformance plan avoids over-claiming while steadily raising technical fidelity.

## Current State (as of March 9, 2026)

- Package/manifest integrity checks are implemented.
- XML parse checks run across core and manifest-declared XML members.
- `content.xml` root/body semantic checks are implemented for text/spreadsheet/presentation mimetypes.
- Optional Relax NG path exists but requires explicit schema mapping.
- ODF reference-corpus and compare tooling exists; baseline generated without wired external validators.

## Goal

Reach a defensible "conformance-core" ODF validation level aligned with OASIS requirements for package,
schema, semantic integrity, and security-relevant package features, while keeping developer workflows practical.

## Non-Goals

- Perfect message-by-message parity with every external ODF validator.
- Reimplementing every validator-specific heuristic outside normative ODF scope.
- Making Java reference validators a hard requirement for local development.

## References

### Normative

- OASIS OpenDocument v1.3 / v1.4 (package model, manifest, document schemas, signatures, encryption model).

### Behavior Calibration Sources

1. ODF Toolkit Validator
2. OPF ODF Validator
3. Official ODF Relax NG schemas

## Decisions

1. Keep ODF validation as a dedicated pipeline (`OdfPackage` + `OdfValidator`), separate from OOXML internals.
2. Expand Python-native conformance first; calibrate against external validators continuously.
3. Keep reference-validator runs optional in local/dev CI, but enforce in scheduled calibration jobs.
4. Preserve current severity model (`ERROR` fatal, `WARNING` non-fatal) with explicit policy docs.
5. Add conformance levels so users can opt into deeper checks deterministically.

## Conformance Levels

1. `foundation`: package/manifest integrity + XML parse + core semantic guards.
2. `schema-core`: Relax NG validation for all in-scope manifest-declared XML members.
3. `semantic-core`: cross-part and content constraints for document class invariants.
4. `security-core`: signature and encryption structure + cryptographic validation policy checks.
5. `reference-calibrated`: tracked drift against ODF Toolkit/OPF on pinned corpus.

## Scope

### In Scope

- Full Relax NG coverage for package-declared XML members (with resolver/version policy).
- ODF semantic rule families for text/spreadsheet/presentation core interoperability constraints.
- Signature and encryption handling policy, diagnostics, and verification hooks.
- Reference parity automation and mismatch taxonomy refinement.
- Documentation that explicitly maps implemented checks to conformance level.

### Out of Scope

- Unlimited format-specific business logic not grounded in ODF specs.
- Proprietary suite-specific extension behavior guarantees.

## Plan

## Phase 0 (March 10-12, 2026): Corpus and Gap Baseline

Workstreams:

1. Expand pinned ODF corpus with targeted conformance edge cases.
2. Wire external validator command templates in calibration environment.
3. Generate baseline mismatch taxonomy with grouped root causes.

Acceptance Criteria:

- Corpus includes valid/invalid examples for each major gap family.
- Baseline reports include real external validator data (not `unavailable`).
- Top mismatch families are categorized by package/schema/semantic/security.

## Phase 1 (March 13-17, 2026): Full Relax NG Coverage

Workstreams:

1. Add schema bundle/version mapping for ODF 1.2/1.3/1.4 variants.
2. Validate all manifest-declared XML members with schema routing rules.
3. Add resolver and include/import handling for official RNG schema sets.

Acceptance Criteria:

- Schema-core mode validates all in-scope XML members deterministically.
- Missing/unresolvable schema cases emit clear diagnostics.
- Test suite includes pass/fail fixtures across ODT/ODS/ODP and auxiliary XML members.

## Phase 2 (March 18-22, 2026): Semantic Core Rules

Workstreams:

1. Define semantic rule registry for ODF families (text, spreadsheet, presentation).
2. Implement high-impact rules:
   - body/document-class alignment beyond `content.xml`
   - manifest media-type/path consistency for key sub-documents
   - cross-part references and required companion-part presence
3. Add deterministic `id` fields for semantic diagnostics.

Acceptance Criteria:

- Semantic-core checks are feature-gated and documented.
- Rule set catches representative interoperability defects with low false-positive rate.
- Tests include positive and negative fixtures for each rule family.

## Phase 3 (March 23-26, 2026): Signature and Encryption Core

Workstreams:

1. Define explicit signature/encryption policy contract.
2. Validate signature package structure (`META-INF/documentsignatures.xml`, related artifacts).
3. Validate encrypted-package structure and emit actionable diagnostics.
4. Add optional cryptographic verification hook path (dependency-gated).

Acceptance Criteria:

- Security-core mode reports signature/encryption structural issues predictably.
- Unsupported crypto states are surfaced as explicit policy diagnostics.
- Security fixtures cover expected valid/invalid layouts.

## Phase 4 (March 27-30, 2026): Reference Calibration Gate

Workstreams:

1. Enrich mismatch normalization for cross-tool comparability.
2. Add calibration CI workflow with threshold policies.
3. Add waiver model for temporary known drift with expiry.

Acceptance Criteria:

- Scheduled calibration job produces reproducible comparison artifacts.
- Drift thresholds are enforced and documented.
- Waiver process is explicit and time-bounded.

## Artifacts to Add or Update

- `src/openxml_audit/odf/package.py`
- `src/openxml_audit/odf/validator.py`
- `src/openxml_audit/odf/` (new schema/semantic/security modules)
- `tests/fixtures/odf/` (expanded conformance corpus)
- `tests/test_odf_package.py`
- `tests/test_odf_validator.py`
- `tests/test_odf_semantic.py`
- `tests/test_odf_security.py`
- `scripts/odf/run_reference_validators.py`
- `scripts/odf/compare_reference_results.py`
- `docs/odf_validation_contract.md`
- `README.md` (conformance level matrix + limitations)

## Risks

1. **Risk:** Full RNG integration introduces heavy schema-resolution complexity.
   - Mitigation: pin schema bundles, cache parse trees, and test resolver behavior explicitly.
2. **Risk:** Semantic rules can create false positives.
   - Mitigation: phased rollout with strict/permissive gating and fixture-driven acceptance.
3. **Risk:** External validator drift obscures regressions.
   - Mitigation: pinned versions + normalized mismatch families + waiver expiry policy.
4. **Risk:** Security checks can imply guarantees not actually provided.
   - Mitigation: explicit policy language distinguishing structural checks vs cryptographic verification.

## Definition of Done

All must be true:

1. Conformance levels are implemented and user-visible.
2. Schema-core mode validates all in-scope XML members with pinned RNG mapping.
3. Semantic-core includes documented, tested rule families for ODT/ODS/ODP.
4. Security-core includes signature/encryption structure validation and policy diagnostics.
5. Reference calibration runs against external validators with tracked drift policy.
6. README and docs clearly state capabilities and remaining limitations.
