# Spec: Auto-Dismiss PPTX / XLSX Repair Dialogs

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.7 — the PowerPoint
and Excel oracles now click through their respective repair dialogs
automatically and Escape any leftover info modals after close-with-save.

## Problem

The 0.6.6 first-baseline run (Spec 022) surfaced operational friction:
when the corpus included files that triggered Excel's "found a problem"
repair dialog or PowerPoint's "we found a problem" sheet, the user had
to click "Yes" / "Repair" manually. The Word oracle (Spec 011) has
auto-dismiss via `click_dialog_button` and an alt-button-label list;
the PowerPoint and Excel oracles inherited only detection
(`find_repair_dialog_text`), not dismissal. Corpus walks of more
than a handful of files become impractical without auto-dismiss —
the user can't sit and click for an hour-long run.

A secondary friction was Excel's *follow-up* info modal — after the
primary repair dialog is dismissed and the workbook is saved, Excel
sometimes shows a "Excel was able to open the file by repairing or
removing the unreadable content" sheet with `View` / `Delete` /
X-close buttons. None of those are obvious affirmatives, so even an
auto-dismiss-the-primary path doesn't clear it. The next file in the
corpus then opens with the leftover modal still on screen.

## Why This Matters

- **Corpus-walk viability.** Phase 2 of Specs 019/020/021 depends on
  walking 50+ files per format. Without auto-dismiss, that's
  hundreds of manual clicks. Auto-dismiss is the precondition for
  scaling oracle data collection.
- **Word-PPTX-XLSX feature parity.** Word has had auto-dismiss since
  Spec 011. The asymmetry was an accident of release sequencing; this
  spec closes it.
- **Smoke-test evidence.** A 0.6.6 baseline run on
  `winsemius.tokens.xlsx` took 61.4 seconds with manual dismissal;
  with auto-dismiss in this release the same file roundtrips in
  5 seconds. ~12× speedup for the corpus walks Phase 2 needs.

## Normative References

- `tools/oracle/word_window.py`'s `click_dialog_button` — pattern
  reference. The PPTX and XLSX implementations mirror its System
  Events button-clicking shape.
- `tools/oracle/word_roundtrip.py`'s `_try_dismiss_repair_dialog`
  — the alt-button-label fallback pattern. Generalized to a
  per-app `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS` constant in this
  spec.
- Spec 020 (PPTX oracle) Phase 2: this work.
- Spec 021 (XLSX oracle) Phase 2: this work.

## Approach

### Phase 1 — auto-dismiss + leftover modal cleanup (this release, 0.6.7)

1. **`pptx.osa.click_dialog_button(label)`** and
   **`xlsx.osa.click_dialog_button(label)`** — System Events button
   click by exact label. Mirrors `word_window.click_dialog_button`.
   The `repeat with X in COLLECTION` idiom is fine in System Events
   namespace (the hang documented for `list_open_*` is specific to
   `tell application "Microsoft <App>"`).

2. **`pptx.osa.dismiss_repair_dialog`** and
   **`xlsx.osa.dismiss_repair_dialog`** — high-level helper that
   detects via `find_repair_dialog_text`, tries each label in
   `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`, returns
   `(was_seen, dialog_text, clicked_label)`. `clicked_label` is None
   when the dialog appeared but no button matched — the orchestrator
   bails with `open_failed` in that case, since Word/Excel can't
   proceed past an undismissed modal.

3. **Per-app accept-button preference lists**
   (`REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`):
   - PPTX: `("Repair", "Yes", "Recover", "OK", "Open")` — PowerPoint's
     primary affirmative for repair sheets is `Repair`; fallbacks
     cover other build versions.
   - XLSX: `("Yes", "Repair", "Recover", "OK", "Open")` — Excel's
     primary recovery dialog is `Yes` / `No`.

4. **`pptx.osa.dismiss_any_leftover_modal`** and
   **`xlsx.osa.dismiss_any_leftover_modal`** — sends Escape (key code
   53) to the frontmost app process. Used as a finalize-after-close
   step to clear secondary info modals that don't have an obvious
   accept button (e.g. Excel's "we made repairs — View / Delete"
   sheet).

5. **Orchestrator wiring** in `tools/oracle/{pptx,xlsx}_repair_oracle.py`:
   - During the open-poll loop: scan for the dialog each iteration;
     if seen, dismiss; if dismissal fails (`clicked is None`), bail
     with `open_failed` and a structured note.
   - After open: one more dismiss pass to catch dialogs that landed
     after the workbook registered.
   - After close-with-save: `dismiss_any_leftover_modal()` to Escape
     any secondary info sheet so the next corpus item starts on a
     clean window.
   - New `RoundtripObservation.repair_dialog_button_clicked` field
     records which label cleared the dialog, for debuggability.

6. **Tests:** the new `dismiss_repair_dialog` / `click_dialog_button`
   / `dismiss_any_leftover_modal` symbols are checked in `__all__`
   contract tests; the dismiss helper's return shape is covered by a
   no-dialog-visible smoke (`(False, None, None)`).

### Phase 2 — fold improvements back into Word's oracle (later)

The Word oracle predates this generalization. Its
`click_dialog_button` and dialog-dismissal logic live in
`tools/oracle/word_window.py` / `tools/oracle/word_roundtrip.py`
rather than in `src/openxml_audit/docx/osa.py`. A future cleanup
release should:

- Lift `click_dialog_button` / `dismiss_repair_dialog` /
  `dismiss_any_leftover_modal` / `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`
  into `docx.osa` for symmetry with PPTX and XLSX.
- Have `word_repair_corpus.py` / `word_roundtrip.py` consume the
  in-package helpers.
- Delete the duplicate primitives from `tools/oracle/word_window.py`.

That cleanup is intentionally scoped out of 0.6.7 — it touches
existing Word oracle behavior that the matrix-driven oracle
(`word_repair_oracle.py`, Spec 010 / 011) depends on, so it
warrants its own release with focused review.

## Acceptance Criteria (Phase 1 / 0.6.7)

1. `pptx.osa` and `xlsx.osa` each export `click_dialog_button`,
   `dismiss_repair_dialog`, `dismiss_any_leftover_modal`, and
   `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS` in `__all__`.
2. The orchestrators in `tools/oracle/{pptx,xlsx}_repair_oracle.py`
   call `dismiss_repair_dialog` interleaved with the open-poll and
   `dismiss_any_leftover_modal` after close-with-save.
3. `RoundtripObservation` for both formats includes
   `repair_dialog_button_clicked: str | None`.
4. Smoke test: `winsemius.tokens.xlsx` (the file that needed manual
   click in 0.6.6) roundtrips end-to-end without manual interaction
   and reports `repair_dialog_button_clicked: "Yes"`.
5. Existing tests pass; new tests cover the dismiss-helper's
   return-shape contract on no-dialog-visible state.
6. CHANGELOG updated. Spec 023 committed.

## Out of Scope (Phase 1)

- Word oracle migration (deferred to Phase 2).
- Per-finding categorization of WHAT Excel/PowerPoint repaired
  (still only "preserved" vs "repaired" outcome — that's Phase 2 of
  Spec 022's per-part text differ).
- Localization. The button labels assume English UI; non-English
  Office builds will need locale-specific label lists. Add as
  observed.
- Behavior on dialogs without a known accept button — currently we
  bail with `open_failed` (after attempting all labels). A future
  release could add a configurable "Escape and continue anyway" mode
  for corpus walks where soft failures are acceptable.

## Risks

- **Button labels drift across Office releases.** Mitigation: the
  preference list is generous (5 labels per app); updates land
  observationally as new label wordings appear in corpus walks.
- **Escape doesn't always close every modal.** Some Office sheets
  bind Escape to "Cancel" which may have side effects (e.g. cancel a
  save). Mitigation: `dismiss_any_leftover_modal` runs *after*
  close-with-save, so cancel-on-Escape can't undo a save that's
  already committed. Pre-close Escape would be unsafe.
- **AppleScript timeouts.** All new helpers catch
  `subprocess.TimeoutExpired` and return False / (False, None, None)
  rather than raising. Cold-launching the target app on first call
  will exceed the 5-10s budgets — that's acceptable; the oracle
  retries on the next poll iteration.
- **Localization.** Non-English Office builds have different button
  labels (e.g. "Sì" for "Yes" in Italian). A globally-deployable
  oracle needs locale-aware label lists; for now we pin to en-US.
