# Spec: ODF Roundtrip Oracle (LibreOffice / soffice)

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.3 — crash-resistant
soffice harness + `observe()` oracle + JSON outcome reporting.

## Problem

The validator's mission is "will this file open in its target app?"
(per `CLAUDE.md`). For ODF, the target apps are LibreOffice and Apache
OpenOffice. We have a strong ODF schema/semantic validator but no way
to elicit "what the target app actually does on this input." The Word
roundtrip oracle (Spec 011) closes that gap for `.docx` against
Microsoft Word; this spec is the ODF sibling.

The signal we want: take an ODF file, ask LibreOffice to open and
re-save it in the same format, and compare the input to the output.
soffice's headless converter (`soffice --convert-to odt ...`) does
exactly that, but is operationally fragile — single-instance lock,
no internal timeout, helper processes that survive SIGKILL of the
parent, font-related crashes, and silent repair behavior (LibreOffice
canonicalizes input on save without surfacing what changed).

## Why This Matters

- **Coverage parity across formats.** Today the oracle ladder is
  Word-only. Specs 010/011 are weeks ahead of equivalent ODF and
  Excel/PowerPoint work. The validator's marketed scope is "OOXML and
  ODF"; oracle ladders should match.
- **Source classification.** Spec 018 (`SourceClass.ODF_NATIVE`)
  routes ODF findings through their own filter. The oracle is the
  empirical signal that converts ODF findings from "schema/semantic
  inferences" into "LibreOffice agrees / disagrees."
- **Spec 013's self-parity gate** wants to filter findings by source
  class. Without an oracle for ODF, every ODF-native finding is
  unverifiable empirically — we'd be promising regression detection
  on signals we can't independently confirm.

## Normative References

- `tools/oracle/word_window.py` and `tools/oracle/word_repair_oracle.py`
  — pattern reference. The ODF oracle mirrors the shape but uses
  soffice rather than osascript-driven Word.
- `tools/parity/dotnet_validator_runner/` — analogous pattern at the
  parity layer (the .NET SDK is to OOXML parity what soffice is to
  ODF roundtrip).
- LibreOffice command-line: `soffice --headless --convert-to <format>`,
  with `-env:UserInstallation=file://<dir>` to avoid the
  single-instance profile lock.
- Spec 011 (Word roundtrip oracle) — sibling.
- Spec 018 (`SourceClass`) — consumer of the oracle's signal.

## Approach

### Phase 1 — soffice harness + minimal oracle (this release, 0.6.3)

Three files:

1. `tools/oracle/odf_window.py` — soffice invocation primitives.
   - `find_soffice()` locates the binary (macOS bundle path first, then
     PATH lookup).
   - `roundtrip(input, output_dir, *, target_format, timeout_seconds)`
     returns a `SofficeRunResult` with one of six outcomes:
     `ok` / `timeout` / `crash` / `missing_output` / `exit_nonzero`.
     Each invocation gets a fresh `UserInstallation` profile dir;
     timeouts kill the process group via `pgrep` + SIGKILL on
     descendants.
   - Public alias `SofficeError` for callers wanting a single sentinel.

2. `tools/oracle/odf_repair_oracle.py` — observation layer.
   - `observe(input_path, work_dir)` runs the harness, fingerprints
     the canonical ODF parts (`content.xml` / `styles.xml` /
     `meta.xml` / `settings.xml`) before and after, and emits a
     `RoundtripObservation` with `outcome ∈ {preserved, repaired,
     crash, timeout, open_failed, missing_output}`.
   - `observe_batch(inputs, work_root)` walks a corpus.
   - CLI: `python tools/oracle/odf_repair_oracle.py FILES... [--output report.json]`.

3. `tests/test_odf_roundtrip_oracle.py` — split into pure-Python
   logic tests (always-on) and soffice-required smoke tests (skipped
   when no soffice binary exists, so CI without LibreOffice still
   passes).

### Phase 2 — corpus expansion + structural diff (0.6.4 candidate)

- TokenMoulds-driven corpus generation. The sibling project's
  `WriterTemplateEmitter` / `CalcTemplateEmitter` /
  `ImpressTemplateEmitter` / `DrawTemplateEmitter` produce all four
  ODF flavors parametrically. Wire in as the oracle's primary corpus.
- Structural diff (XML-level, not byte-level) of `content.xml` parts —
  distinguishes cosmetic repairs (attribute reorder, formatting reflow)
  from substantive ones (element added/removed/reordered).
- LibreOffice qa import as a secondary corpus (real-world wildness).

### Phase 3 — scenario generator + categorical baselines (0.6.5+)

Sibling of `tools/oracle/scenarios/property_ordering.py` for ODF.
Pairwise mutations of property elements; commit JSON baselines under
`tools/oracle/baselines/odf_*.json`.

## Acceptance Criteria (Phase 1 / 0.6.3)

1. `find_soffice()` locates the binary on macOS without explicit
   configuration.
2. `roundtrip()` survives a 60-second hung soffice invocation by
   force-killing the process group (verified by integration test;
   too time-consuming to exercise in unit tests, manually validated
   for now).
3. The oracle classifies a real-world TokenMoulds-emitted ODT as
   either `preserved` or `repaired` (not `crash` / `timeout` /
   `open_failed`) on a known-good input. Verified manually before
   release.
4. Tests pass on systems without soffice (skip pattern).
5. Tests pass on systems with soffice (the 8 in
   `test_odf_roundtrip_oracle.py`).
6. CHANGELOG updated. Spec 019 committed.

## Out of Scope (Phase 1)

- Structural XML diff (Phase 2).
- Corpus generation via TokenMoulds (Phase 2).
- Pairwise scenario mutations (Phase 3).
- LibreOffice qa import (Phase 2 secondary, may slip).
- Wiring the oracle's signal into a CI gate (Spec 013's job, not this).
- macOS-specific notarization / sandbox-aware paths beyond what
  soffice already handles.

## Risks

- **soffice version drift.** macOS LibreOffice 26.x is what the
  current dev machine runs. Older versions (e.g., LibreOffice on
  Linux CI) may have different `--user-profile` syntax, different
  default fonts, different repair behaviors. Mitigation: oracle
  records the soffice version with each observation (future work).
- **Corpus walks at scale will hit crashes the harness has never
  seen.** Phase 2's first 1000-file run will surface failure modes
  the harness's `outcome` enum doesn't yet name. Plan for adding
  cases as they appear.
- **TokenMoulds output may not be diverse enough.** A single
  emitter family produces files in one canonical shape; soffice's
  repair behavior on broken-by-other-emitters input (e.g.
  Word-saved-as-ODT, Pages-exported, Google-Docs-exported) is
  invisible. LibreOffice qa import (Phase 2 secondary) is the
  hedge.
- **Headless soffice is silent about repairs.** Unlike Word's modal
  "unreadable content" dialog, LibreOffice canonicalizes on save
  without warning. The oracle infers repairs from the diff, which
  loses the categorical signal Word's oracle provides. Phase 2's
  structural diff narrows this gap; full parity requires running
  soffice non-headless to capture UI-level warnings (deferred).
