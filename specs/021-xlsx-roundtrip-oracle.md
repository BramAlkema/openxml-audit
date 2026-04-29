# Spec: Excel Roundtrip Oracle

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.5 — completes the
roundtrip oracle ladder across all four formats (Word + ODF + PPTX +
**Excel**).

## Problem

Excel was the most-neglected format in the validator: no
roundtrip oracle, no integration tests touching Excel-specific
behaviors, no per-format spec since 002 (which was generic parity
work). The CHANGELOG entry up through 0.6.4 lists Word, PowerPoint,
and ODF oracle ladders; Excel's silence was conspicuous.

The validator's stated mission is roundtrip survival in the target
app. With Word ✓ (Spec 011), ODF ✓ (Spec 019), PowerPoint ✓ (Spec
020), Excel is the last format without a corresponding oracle. This
spec closes that gap.

## Why This Matters

- **Format parity.** The validator's marketed scope is "OOXML and
  ODF". Shipping three of four formats with roundtrip oracles and
  the fourth with none is incoherent.
- **Source-class signal.** Spec 018 added
  `SourceClass.EXCEL_APP_COMPAT` for Excel-specific findings beyond
  what the .NET SDK reports. Without an Excel oracle, every such
  finding remains an inference, not an observation.
- **Spec 013's eventual oracle gate** wants oracle observations
  across all four formats. Excel was the only remaining hole.

## Normative References

- `tools/oracle/word_repair_oracle.py` (Spec 011) — pattern
  reference; also relies on the same `~/Documents/.<app>_oracle_runs`
  staging convention to navigate App Sandbox.
- `tools/oracle/odf_repair_oracle.py` (Spec 019) — closest
  pattern reference; both use hash-based per-part fingerprint diff
  in the absence of a per-format `lab` differ.
- `tools/oracle/pptx_repair_oracle.py` (Spec 020) — closest
  shape reference; same `RoundtripObservation` contract.
- `src/openxml_audit/xlsx/osa.py` — existing Excel window primitives
  this spec extends.
- `src/openxml_audit/docx/osa.py`, `src/openxml_audit/pptx/osa.py`
  — pattern references for the new `close_*_saving` /
  `is_*_open` / repair-dialog primitives.
- `tools/oracle/preflight.py` — generalized in this release to
  cover all four engines; documents the macOS permission model.
- `docs/oracle_permissions.md` — new in this release; the macOS
  Automation / Accessibility setup checklist.

## Approach

### Phase 1 — extend xlsx/osa, build orchestrator (this release, 0.6.5)

1. **Extend `src/openxml_audit/xlsx/osa.py`** with the primitives
   that exist in `docx.osa` and `pptx.osa` but were missing in
   `xlsx.osa`:
   - `list_open_workbook_names()` — query Excel's `workbooks`
     collection. Catches `subprocess.TimeoutExpired` so that a
     not-running Excel doesn't hang the test suite.
   - `is_workbook_open(path)` — predicate over above.
   - `close_workbook_saving()` — close active with `saving yes`
     (the proven persist path from Word + PowerPoint).
   - `find_repair_dialog_text()` — System Events scan for Excel's
     "found a problem" / "could not be opened" modal.
   - `REPAIR_DIALOG_PATTERNS` — public match list.

2. **Build `tools/oracle/xlsx_repair_oracle.py`** as the orchestrator.
   Stages input under `~/Documents/.xlsx_oracle_runs/<id>/` (Excel
   App Sandbox default), uses existing `xlsx.osa.open_workbook`,
   polls `is_workbook_open`, detects repair dialog, calls
   `close_workbook_saving`. Diffs with a hash-based per-part
   fingerprint over the canonical OOXML parts (workbook.xml,
   sheets, styles, etc. — see `_FINGERPRINTED_PREFIXES`).

3. **Generalize `tools/oracle/preflight.py`** to cover Word + Excel
   + PowerPoint + LibreOffice. The original Word-only `check()` is
   preserved as a back-compat alias for `check_word()`. Run with
   `python -m tools.oracle.preflight` (all engines) or
   `--engine <name>` for one.

4. **Document the macOS setup** in `docs/oracle_permissions.md`.
   Covers Automation (control of Word/Excel/PowerPoint),
   Accessibility (System Events keystrokes for `Cmd-S` save paths),
   sandbox staging directories, and the symptom checklist for
   denied/revoked permissions.

5. **`tests/test_xlsx_roundtrip_oracle.py`** — 11 tests: 9 always-on
   (harness logic, fingerprinting, summary shape, env-var override,
   `__all__` contract) + 2 Excel-required smoke tests that skip
   cleanly when Excel isn't installed.

### Phase 2 — corpus expansion + repair categorization (later)

- TokenMoulds-driven Excel corpus generation. The sibling project
  has full XLSX/XLTX emitters
  (`emitters/excel.py`, `emitters/excel_stylesheet.py`) and three
  pre-generated `.xltx` files in `generated/`. Wire as the
  oracle's primary corpus.
- Lift `pptx.lab.compare_pptx_packages`'s diff machinery into a
  shared `openxml_audit.parts.diff` so the Excel oracle can
  produce per-part text diffs (currently only hash deltas).
- Pairwise mutation scenarios for Excel's most-repaired property
  elements (TBD by the corpus walk).
- Auto-dismiss the repair dialog (Word oracle pattern; Phase 2
  for both PPTX and XLSX).

## Acceptance Criteria (Phase 1 / 0.6.5)

1. `xlsx.osa` exposes the six new primitives + `REPAIR_DIALOG_PATTERNS`.
2. `tools/oracle/xlsx_repair_oracle.py` orchestrates a roundtrip via
   the in-package `xlsx.osa` layer — no duplicate primitives.
3. `tools/oracle/preflight.py` covers all four engines and is
   back-compat with the existing Word-only `check()` API.
4. `docs/oracle_permissions.md` documents the macOS Automation +
   Accessibility setup with concrete System Settings paths.
5. Tests pass on systems without Excel (skip pattern).
6. Tests pass on systems with Excel.
7. CHANGELOG updated. Spec 021 committed.

## Out of Scope (Phase 1)

- Auto-dismiss the repair dialog (Phase 2; same as PPTX/Word).
- Corpus generation (Phase 2; relies on TokenMoulds wiring).
- Per-part text diffs (Phase 2; currently hash-only).
- Pairwise mutation scenarios (Phase 2).
- Refactoring `tools/oracle/word_window.py` to use the in-package
  `osa` layer (separate cleanup; the duplication exists and is
  outside this spec's scope).
- CI wiring (Spec 013's job).

## Risks

- **Permission state drift.** Office app updates and macOS minor
  releases occasionally drop Automation grants. Mitigation:
  preflight + `docs/oracle_permissions.md`.
- **Excel's `save active workbook` is sometimes flaky between
  builds.** The close-with-save path is the proven workaround
  (same finding as Word's oracle). Mitigation: existing
  `close_workbook_saving` primitive routes through the close
  handler, which has been stable across Office 2016+.
- **App Sandbox staging path.** Same `~/Documents` constraint as
  Word and PowerPoint.
- **Repair-dialog detection.** Excel's modal text varies more than
  Word's or PowerPoint's (different wording for cell-corruption,
  shared-strings repair, formula-bound repair). The pattern list
  is generous; corpus walks may surface dialog wordings the list
  doesn't catch — add as observed.
- **Phase 1's hash-based diff is coarse** — it can't distinguish
  "Excel reordered an attribute" (cosmetic) from "Excel deleted a
  shared-string entry" (substantive). Phase 2's per-part text
  diff narrows this; for Phase 1 the binary preserved/repaired
  signal is enough to bootstrap the corpus.
