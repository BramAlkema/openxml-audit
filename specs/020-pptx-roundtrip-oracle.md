# Spec: PowerPoint Roundtrip Oracle

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.4 — completes the
existing PPTX oracle layer by adding the missing close-with-save +
repair-dialog primitives and an orchestrator that wires
`pptx.osa` and `pptx.lab.compare_pptx_packages` into the oracle
observation shape used by Word and ODF.

## Problem

The validator's stated mission is roundtrip survival in the target
app. For Word we have `tools/oracle/word_repair_oracle.py` (Spec 011)
and for ODF we have `tools/oracle/odf_repair_oracle.py` (Spec 019).
PowerPoint is the third OOXML target — and the project has substantial
prior PPTX tooling already on disk:

| What exists | Where | Role |
|---|---|---|
| Generic osascript primitives | `src/openxml_audit/osa/__init__.py` | `osascript`, `applescript_quote`, `launch_app`, JXA helpers |
| PowerPoint window primitives | `src/openxml_audit/pptx/osa.py` | `launch_powerpoint`, `open_presentation_via_ui`, window discovery |
| PPTX timing oracle | `src/openxml_audit/pptx/oracle.py` | extracts/normalizes timing trees from a deck |
| **PPTX package differ** | `src/openxml_audit/pptx/lab.py` | `compare_pptx_packages` — full per-part snapshot + diff with timing-change rollup |
| Oracle starter decks | `pptx/oracle_starter_deck.py`, `pptx/timing_oracle_deck.py` | input fixture generators |

What's missing is the **persist + diff orchestrator**: the file that
opens a PPTX in PowerPoint, closes-with-save, and produces a
`RoundtripObservation` matching the Word and ODF shapes. The
prerequisite `save_presentation` / `close_presentation_saving` /
repair-dialog primitives are also missing from `pptx.osa` (the
analogous Word and Excel modules already have them — see
`src/openxml_audit/docx/osa.py` and `src/openxml_audit/xlsx/osa.py`).

## Why This Matters

- **Coverage parity across formats.** Word ✓ (Spec 011), ODF ✓ (Spec
  019), PowerPoint = next. Excel = 0.6.5.
- **Source-class signal for PPTX findings.** Spec 018 added
  `SourceClass.POWERPOINT_APP_COMPAT`. Without a PowerPoint roundtrip
  oracle, every PowerPoint-app-compat finding (e.g. presProps /
  viewProps / tableStyles missing-relationship repair) is unverifiable
  empirically.
- **Spec 013's eventual oracle gate** wants oracle observations across
  all four formats. PPTX is the format with the most existing
  scaffolding (timing oracle, deck generators, package differ); not
  closing the gap leaves the rest of the ladder unbalanced.

## Normative References

- `tools/oracle/word_repair_oracle.py` + `tools/oracle/word_window.py`
  — Word sibling.
- `tools/oracle/odf_repair_oracle.py` + `tools/oracle/odf_window.py`
  — ODF sibling.
- `src/openxml_audit/pptx/osa.py` — existing PowerPoint window
  primitives this spec extends.
- `src/openxml_audit/pptx/lab.py` — `compare_pptx_packages` — the
  per-part diff this spec hands the oracle's outputs to.
- `src/openxml_audit/docx/osa.py`, `src/openxml_audit/xlsx/osa.py`
  — pattern reference for the new `save` / `close` primitives.

## Approach

### Phase 1 — extend pptx/osa, build orchestrator (this release, 0.6.4)

1. **Extend `src/openxml_audit/pptx/osa.py`** with primitives that are
   present in the docx/xlsx siblings but missing here:
   - `list_open_presentation_names()` — query PowerPoint's
     `presentations` collection.
   - `is_presentation_open(path)` — predicate over above.
   - `save_presentation()` — Cmd-S keystroke save (mirrors
     `docx.osa.save_document`).
   - `close_presentation()` — close active without saving.
   - `close_presentation_saving()` — close active with `saving yes`
     (the proven-reliable persist path; same shape as Word's oracle).
   - `find_repair_dialog_text()` — System Events scan for
     PowerPoint's "found a problem" / "repair" modal.
   - `REPAIR_DIALOG_PATTERNS` — public match list.

2. **Build `tools/oracle/pptx_repair_oracle.py`** as a thin
   orchestrator. Stages input under
   `~/Documents/.pptx_oracle_runs/<id>/` (PowerPoint App Sandbox
   default), opens via existing `open_presentation_via_ui`, polls
   `is_presentation_open`, detects (without dismissing) any repair
   dialog, calls `close_presentation_saving`, then hands the original
   + post-PowerPoint copy to `compare_pptx_packages` for the per-part
   diff. Returns a `RoundtripObservation` matching Word/ODF shape.

3. **`tests/test_pptx_roundtrip_oracle.py`** — 5 always-on tests for
   the harness logic + 2 PowerPoint-required smoke tests that skip
   cleanly when PowerPoint isn't installed (so CI without it passes).

### Phase 2 — corpus expansion + repair categorization (0.6.5+ candidate)

- TokenMoulds-driven PPTX corpus generation (the
  `ImpressTemplateEmitter` for ODP plus existing PPTX template
  emitters).
- Repair categorization on top of `compare_pptx_packages`'s already-
  detailed per-part diff output: distinguish cosmetic (XML reflow)
  from substantive (slide content change, layout reflow).
- Pairwise scenario mutations on PowerPoint property elements —
  analogous to the Word oracle's CT_TrPr/CT_TblPr/CT_TcPr/CT_SectPr
  baselines.

## Acceptance Criteria (Phase 1 / 0.6.4)

1. `pptx.osa` exposes the six new primitives + `REPAIR_DIALOG_PATTERNS`.
2. `tools/oracle/pptx_repair_oracle.py` orchestrates a roundtrip via
   the in-package `pptx.osa` and `pptx.lab.compare_pptx_packages`
   layers — no duplicate primitives.
3. Tests pass on systems without PowerPoint (skip pattern).
4. Tests pass on systems with PowerPoint.
5. CHANGELOG + spec committed.

## Out of Scope (Phase 1)

- Auto-dismissing the repair dialog (Phase 2; Word oracle has this,
  ODF doesn't need it because soffice is silent).
- Corpus generation (Phase 2).
- Repair categorization (cosmetic vs substantive — Phase 2).
- Pairwise mutation scenarios (Phase 2).
- Refactoring `tools/oracle/word_window.py` to use the in-package
  `osa` layer (separate cleanup; the duplication exists and is
  outside this spec's scope).
- CI wiring (Spec 013's job).

## Risks

- **PowerPoint UI changes.** AppleScript dictionary stability for
  `presentations` / `active presentation` / `close ... saving yes`
  has held since Office 2016 but isn't guaranteed forever. Mitigation:
  the oracle records the PowerPoint version with each observation in
  the future (Phase 2; see Word oracle's `word_version()`).
- **App Sandbox staging path.** PowerPoint may revoke
  `~/Documents` access in a future macOS version. Mitigation:
  `PPTX_ORACLE_STAGE` env var lets the user override.
- **Sibling-project drift.** `src/openxml_audit/pptx/lab.py`'s
  `compare_pptx_packages` is shared with other tooling. Changes
  there could surprise the oracle. Mitigation: tests pin the report
  shape (Phase 2 work; Phase 1 trusts the existing contract).
- **PowerPoint stealing focus.** Roundtrip runs visibly hijack the
  user's screen during corpus walks. Mitigation: document this as a
  dev-machine constraint; long runs should happen on a dedicated
  machine. Same constraint as the Word oracle.
