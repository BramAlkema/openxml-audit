# Spec: Word Oracle In-Package OSA Consolidation

## Status

Proposed (April 30, 2026). Phase 1 of Spec 026's roadmap to 0.8.0
that was deferred from 0.7.3. Folds `tools/oracle/word_window.py`'s
primitives into `openxml_audit.docx.osa` so the four-format oracle
ladder uses one consistent in-package osa layer.

## Problem

The Word oracle (Spec 011, 0.5.0) shipped its window primitives in
`tools/oracle/word_window.py` because that's where roundtrip-oracle
infrastructure lived at the time. When PPTX (Spec 020, 0.6.4) and
XLSX (Spec 021, 0.6.5) added their oracles, they used the in-package
layer (`src/openxml_audit/<format>/osa.py`) for symmetry with the
existing `docx.osa` / `xlsx.osa` modules and to leverage the shared
`openxml_audit.osa` primitives. The result through 0.7.3 was an
asymmetric layout:

| Format | Window primitives | Repair dialog handling |
|---|---|---|
| Word    | `tools/oracle/word_window.py` (16 functions) | here |
| PPTX    | `src/openxml_audit/pptx/osa.py` (consolidated 0.6.4 + 0.6.7) | here |
| XLSX    | `src/openxml_audit/xlsx/osa.py` (consolidated 0.6.5 + 0.6.7) | here |
| ODF     | `tools/oracle/odf_window.py` (soffice headless, separate shape) | n/a |

Word's primitives weren't reachable from in-package code without
reaching into `tools/`. That's been a known wart since 0.6.7's
auto-dismiss release; Spec 023 deferred this consolidation
explicitly to keep the auto-dismiss work focused on PPTX and XLSX.

This spec resolves the asymmetry.

## Why This Matters

- **One consistent in-package osa layer.** `docx.osa` / `pptx.osa` /
  `xlsx.osa` now have the same shape and symbol set. Future
  features (e.g. localization of dialog patterns, schema-version
  detection from app version, etc.) ship once and reach all three.
- **Word primitives become consumable from validator code.** A
  future canonical-form check (parallel to Spec 029's
  `Excel_InlineStrCells`) that wants to detect "Word will rewrite
  this on save" can import `docx.osa` directly. Reaching into
  `tools/oracle/` from `src/openxml_audit/` is a layering smell;
  this fixes it.
- **Spec 011 / Spec 010's matrix-driven oracle still works.**
  `tools/oracle/word_window.py` becomes a back-compat shim that
  re-exports from the in-package layer. No consumer changes
  required; the consolidation is mechanically invisible to existing
  callers.

## Normative References

- `specs/011-word-roundtrip-oracle.md` — original Word oracle
  surface; defined `tools/oracle/word_window.py`.
- `specs/020-pptx-roundtrip-oracle.md` — pattern reference: PPTX
  consolidated into `pptx.osa` directly.
- `specs/021-xlsx-roundtrip-oracle.md` — same shape for XLSX.
- `specs/023-auto-dismiss-pptx-xlsx-repair-dialogs.md` — Phase 2
  note explicitly deferring the Word consolidation to a focused
  release. That release is this one.
- `specs/026-self-parity-sovereign-gate-roadmap.md` — umbrella
  for the path to 0.8.0; Word consolidation listed as 0.7.3 then
  slipped to 0.7.4 after the Excel canonical-form pivot.

## Approach

### Phase 1 — consolidate (this release, 0.7.4)

1. **Extend `src/openxml_audit/docx/osa.py`** with every primitive
   `tools/oracle/word_window.py` exposed:
   - Constants: `WORD_APP_BUNDLE`, `WORD_APP_ID`,
     `WORD_PROCESS_NAME`, `REPAIR_DIALOG_PATTERNS`,
     `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`.
   - App lifecycle: `launch_word`, `launch_word_app` (alias),
     `is_word_running`, `word_version`.
   - Document operations: `open_document`,
     `list_open_document_names`, `is_document_open`,
     `activate_document`.
   - Save/close: `save_document`, `close_document` (no save),
     `close_document_saving` (with save), and historical aliases
     `close_active_document` / `close_active_document_saving`.
   - Dialog handling: `find_repair_dialog_text`,
     `click_dialog_button`, `dismiss_repair_dialog`,
     `dismiss_any_leftover_modal`. Mirrors the PPTX/XLSX
     auto-dismiss pattern from 0.6.7.

2. **Convert `tools/oracle/word_window.py` to a back-compat
   shim**. Re-exports every name from `openxml_audit.docx.osa`
   plus `osascript` / `osascript_jxa` / `applescript_quote` from
   `openxml_audit.osa`. Aliases `OsascriptError = RuntimeError` so
   existing `except word_window.OsascriptError` clauses keep
   catching what they always caught (`openxml_audit.osa.osascript`
   raises plain `RuntimeError`).

3. **`tests/test_docx_osa.py`** — 11 new tests:
   - 4 lock the symbol set + callability + alias identity
     (the contract that consumers depend on).
   - 4 verify the dialog-handling helpers' shape (patterns
     non-empty, accept-button list includes "Open and Repair",
     dismiss helper returns `(False, None, None)` cleanly when
     no dialog visible).
   - 3 verify the back-compat shim re-exports correctly,
     `OsascriptError` is a usable alias, and the internal
     `_applescript_quote` is reachable for legacy callers.

### Future cleanup (deferred)

- **Migrate `tools/oracle/word_roundtrip.py` to import from
  `docx.osa` directly** (drop the `from tools.oracle import
  word_window` indirection). Mechanical change; deferred to keep
  this release tight.
- **Migrate `tools/oracle/word_repair_corpus.py` and
  `word_repair_oracle.py`** to import from `docx.osa` directly.
- **Eventually delete `tools/oracle/word_window.py`** once all
  consumers migrate. Pre-1.0 timeline.

## Acceptance Criteria (Phase 1 / 0.7.4)

1. `src/openxml_audit/docx/osa.py` exposes the 22 symbols listed
   under "App lifecycle / Document operations / Save/close /
   Dialog handling" above.
2. `tools/oracle/word_window.py` is a back-compat shim that
   re-exports those symbols + `OsascriptError` alias +
   `osascript`/`osascript_jxa`/`applescript_quote` from the
   shared `openxml_audit.osa` module.
3. `tests/test_docx_osa.py` passes (11 tests).
4. The full test suite passes (`tests/`, with the 2 ODF
   integration tests still skipped per their environmental
   flakiness).
5. `tools/oracle/word_repair_oracle.py` (the matrix-driven
   oracle from Spec 010/011) continues to work through the
   shim.
6. `tools/oracle/word_repair_corpus.py` (the corpus walker
   from 0.6.6) continues to work through the shim.
7. CHANGELOG updated. Spec 030 committed.

## Out of Scope (Phase 1)

- **Direct migration of consumers** (`word_roundtrip.py`,
  `word_repair_corpus.py`, `word_repair_oracle.py`) to import
  from `docx.osa`. Deferred to keep this release mechanically
  invisible to the matrix oracle. Future work.
- **Deletion of `tools/oracle/word_window.py`**. Same reason.
- **Localization** of `REPAIR_DIALOG_PATTERNS` /
  `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`. The current lists assume
  English Office for Mac M365; non-English builds need locale-
  specific lists. Cataloged as future work in `pptx.osa` /
  `xlsx.osa`'s docstrings; same applies here.
- **The `find_repair_dialog_text` System Events
  scripting** — kept identical to the 0.6.6 implementation.
  Subject to Word UI changes; same caveat as PPTX/XLSX.

## Risks

- **Shim transparency.** If `word_window.<symbol>` was used as
  a sentinel for `is None` checks anywhere (i.e. someone tested
  whether `word_window.list_open_document_names` was None to
  decide whether the module was usable), the shim's re-export
  makes the symbol always non-None. No such usage found in the
  current codebase, but external callers could break.
  Mitigation: `tools/oracle/word_window.py` is internal tooling,
  not a public API; the shim's contract is "every name keeps
  resolving to the same callable."
- **`OsascriptError = RuntimeError` alias is too broad.**
  Historically, `word_window.osascript` raised
  `OsascriptError` (a RuntimeError subclass) on osascript-level
  failures. Other code paths could raise plain RuntimeError for
  unrelated reasons. The shim's alias catches both. Mitigation:
  no consumer was using the subclass for narrow filtering; the
  alias is functionally equivalent in practice.
- **Future divergence between the four osa modules.** If
  `pptx.osa` evolves a feature `docx.osa` doesn't (or vice
  versa), the symmetry advertised by this consolidation gets
  uneven. Mitigation: each osa module's `__all__` makes the
  divergence visible; tests like
  `test_docx_osa_exports_full_word_primitive_set` lock the
  contract.
