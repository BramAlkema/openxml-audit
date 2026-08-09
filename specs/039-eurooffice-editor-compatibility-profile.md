# Spec 039: EuroOffice Editor Compatibility Profile

## Status

Implemented (August 9, 2026): an additive, version-bound classifier for the
known static OOXML drift observed after a real Nextcloud EuroOffice editor
callback. The Python API and CLI JSON report preserve the raw strict result.

## Problem

The conversion oracle in Spec 038 does not exercise the Nextcloud editor save
callback. A real callback can produce a document that EuroOffice opens while
strict validation reports deterministic serializer findings. Treating every
such finding as a new failure makes a library checker noisy; globally ignoring
the finding families would hide corruption, real active content, or drift after
an upgrade.

The checker therefore needs two simultaneous answers:

1. the unmodified strict validation result;
2. whether every finding fits one exact, observed EuroOffice environment.

Compatibility is not schema validity. An accepted compatibility result must
never rewrite `ValidationResult`, remove findings, or claim that content and
semantics survived the editor roundtrip.

## Profile Identity and Evidence Boundary

The first profile is:

```text
tokenooxml-eurooffice-editor-9.3.1.37-connector-11.0.1
```

It is bound to:

- live Document Server version `9.3.1.37`;
- Nextcloud EuroOffice connector `11.0.1`;
- strict, unlimited validation with OOXML security checks enabled;
- Microsoft 365 and Office 2007 validation targets;
- TokenOOXML DOCX, PPTX, and XLSX files saved through the real Nextcloud
  editor callback on August 9, 2026.

This is a callback-artifact profile, not a universal claim about all documents
produced by EuroOffice. A new server or connector version requires a new
profile and new callback evidence.

The report always emits:

```json
"semantic_preservation": "not-assessed"
```

Template provenance, template SHA, content/semantic comparisons, and visual
fidelity remain separate quality/compliance checks. A future worker must fail
semantic loss even when the static compatibility profile accepts all schema
findings.

## Classification Contract

`classify_eurooffice_compatibility()` returns one of four statuses:

| Status | Meaning |
|---|---|
| `strict-clean` | The matched environment produced no findings. |
| `accepted-known-drift` | Every finding exactly matched a known rule and stayed within its observed cap. |
| `unexpected-drift` | A finding was unknown, a cap was exceeded, or a security prerequisite failed. |
| `unverified-environment` | Server, connector, file type, validation target, or scan contract did not match the profile. |

Matching uses the normalized error ID plus exact severity, error type, source
class, part URI, node, and description. Rules require either an exact stable
path or the calibrated indexed path shape. Similar wording or a finding in a
different part or structure is not enough.

Occurrence caps are conservative upper bounds from the callback snapshot. They
prevent a broad family match from accepting an unbounded rewrite. The Word
style rules have high caps because EuroOffice expanded a large built-in table
style set; a count above the observed bound is new drift.

## Active-Content Boundary

The callback wrote an OLE `.bin` default content-type declaration without a
matching `.bin` package member. That exact declaration can be classified as
known drift only when the checker can open the package and prove no `.bin`
member exists.

The finding remains unexpected when:

- any `.bin` payload is present;
- the package cannot be inspected;
- the content type, extension, part, node, or finding ID differs;
- more than one matching declaration appears.

This rule does not waive embedded objects or active content.

## CLI Contract

The profile can be added to text, JSON, or XML output:

```bash
openxml-audit callback.docx \
  --output json \
  --eurooffice-profile \
    tokenooxml-eurooffice-editor-9.3.1.37-connector-11.0.1 \
  --eurooffice-document-server-version 9.3.1.37 \
  --eurooffice-connector-version 11.0.1
```

Profile mode defaults the validation target to Microsoft 365 and forces a
strict, unlimited security scan. The existing `valid` field, full finding list,
and strict process exit status are unchanged. JSON adds an
`eurooffice_compatibility` object for consumers such as a Nextcloud worker.

The compatibility JSON is compact: accepted findings are aggregated by rule;
unexpected findings are grouped without discarding the raw validator list.

## Acceptance

- Known, exactly matched findings are accepted without mutating the raw result.
- Unknown findings and near-matches remain unexpected.
- Counts above an observed cap remain unexpected.
- A phantom OLE declaration is accepted only after ZIP-member inspection.
- Real `.bin` payloads and unreadable packages remain unexpected.
- Unknown server, connector, and validation versions accept no findings.
- CLI profile mode forces unlimited collection and security validation.
- CLI JSON retains raw strict validity/counts and labels semantic preservation
  as not assessed.
- The three live callback artifacts reproduce `accepted-known-drift` while
  retaining raw results of DOCX `3107 errors / 4 warnings`, PPTX `1 error`, and
  XLSX `4 errors` in the current source tree.
