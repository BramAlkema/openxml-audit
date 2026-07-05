# Spec: LibreOffice OOXML Roundtrip Oracle — the CI-Runnable Rung

## Status

Proposed (July 5, 2026). Phase 1: oracle tool + feature-survival
probes + first committed baseline (local run) + advisory CI workflow.

## Problem

All three Microsoft-app oracles need a macOS machine with Office and
hand-granted AppleScript permissions — they cannot run in CI, so the
evidence ladder only climbs when someone runs them by hand. The ODF
oracle already proved the alternative: headless soffice is silent,
containerizable, and crash-supervised (`tools/oracle/odf_window.py`).

LibreOffice also opens OOXML. A soffice roundtrip of `.docx`/`.xlsx`/
`.pptx` is real app evidence for a real target app (LibreOffice is a
first-class consumer of OOXML files in the wild), and it is the only
app oracle that can run on every push.

Spec 034 sharpened what to ask it: the generated reference documents
carry one proven feature per slide/section with a manifest mapping
feature → location. So the oracle can report not just "N parts
changed" but **which reference features survived the roundtrip** —
the ladder cycle operating against a live app, in CI.

## Design

### Harness

`odf_window.roundtrip` gains `docx` / `xlsx` / `pptx` target formats
(soffice's `--convert-to` handles them natively; supervision,
ephemeral profiles, and timeout reaping are format-agnostic).

### Oracle (`tools/oracle/lo_ooxml_repair_oracle.py`)

Mirrors `odf_repair_oracle.py` — same `RoundtripObservation` schema,
same summary block, same exit-code contract (`repaired` is data, not
failure). Differences:

- canonical-part filter per OOXML format (document/styles for DOCX;
  workbook/worksheets/sharedStrings/styles for XLSX;
  presentation/slides for PPTX)
- expectation note: LibreOffice virtually always rewrites foreign
  formats byte-level, so `preserved` is rare by construction; the
  signal is open-success, part add/remove, and feature survival.

### Feature-survival probes (`openxml_audit.reference.feature_probes`)

Namespace-aware XPath signatures per capability key (fade/wipe
`animEffect` filters, `endCondLst` time/click conditions, `repeatDur`,
non-root `restart`). Given a roundtripped PPTX and a reference
manifest, the oracle reports per-feature: slide part still present,
signature still present. This is evidence about LibreOffice —
distinct from (and not a substitute for) the PowerPoint tier ladder;
registry tiers are unaffected.

### Baseline + CI

- First baseline committed under `tools/oracle/baselines/lo_ooxml/`
  from a local run over the TokenMoulds v0.7.2 OOXML corpus plus the
  three generated reference documents.
- New workflow `.github/workflows/reference-oracle.yml` (advisory):
  install LibreOffice on ubuntu, build the reference documents at
  `loadable`, run the oracle, upload the observation report as an
  artifact. Non-blocking — it is an evidence collector, not a gate
  (gating stays with Spec 013/026).

## Non-Goals

- No claims about PowerPoint/Word/Excel behavior from LibreOffice
  observations; the report is explicitly LibreOffice-scoped.
- No tier promotion from this oracle (tiers are defined against the
  file's target app).
- No semantic XML diffing beyond part-level change lists and feature
  signatures (cosmetic-vs-substantive categorization stays future
  work, as in Spec 019 Phase 2).

## Acceptance

- Oracle runs locally over the corpus + reference docs; baseline JSON
  committed with a README section describing findings.
- Reference-manifest mode reports per-feature survival for the
  generated deck.
- Unit tests cover format routing, canonical-part filters, and probe
  signatures without requiring soffice; an integration smoke is
  skipped when soffice is absent.
- CI workflow builds references and uploads an observation artifact;
  it does not block merges.
