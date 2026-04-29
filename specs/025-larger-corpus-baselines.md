# Spec: Larger Corpus Baselines (0.6.6 second run)

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.9 — second baseline
run across all four formats with a 4× larger corpus, made tractable
by Spec 023's auto-dismiss (0.6.7) and annotated with the per-part
diffs from Spec 024 (0.6.8). Surfaces real operational signal AND
exposes a corpus-curation gap that becomes the headline limitation
of this release.

## Problem

The 0.6.6 first baseline (Spec 022) ran on 9 files across 4 formats —
enough to prove the tooling works, not enough to characterize
behavior. Phase 2 of Spec 022 deferred broader corpus walks to
"once auto-dismiss lands so they're tractable." With 0.6.7's
auto-dismiss for PPTX/XLSX shipped, and 0.6.8's per-part diff
delivering richer per-file evidence, the precondition for a larger
walk is met.

The new walk surfaces two things at once:

1. **Operational gaps in auto-dismiss** that the smaller first
   baseline didn't expose: dialog wording variants the pattern lists
   don't cover, button labels outside `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`,
   PowerPoint's secondary "[Repaired]" info modal, mid-run sandbox
   fragility.
2. **A corpus-curation problem**: my expanded corpus assembly was
   indiscriminate — I globbed `tokenmoulds/{reports/visual,
   generated/word-demo, scratch/odf,...}` without distinguishing
   release-quality TokenMoulds emitter output from troubleshooting-
   session leftovers. The user flagged this mid-run; it's now a
   documented limitation of this baseline, not a flaw in the
   tooling.

## Why This Matters

- **The first finding (operational gaps) is gold for tooling
  improvements.** Each gap becomes a pattern-list addition or an
  auto-dismiss-flow refinement in a future release. Without the
  larger walk, we don't see them.
- **The second finding (corpus problem) is honest documentation.**
  Cleaning up "where to source production-quality TokenMoulds output"
  is its own work — running TokenMoulds via its Python API to emit
  `.docx`/`.xlsx`/`.pptx` (not templates) needs more infra than
  shows up in `examples/build_acme_*.py`. Punting that to a later
  release with explicit corpus-curation work is more honest than
  pretending the current baseline is comprehensive.
- **Spec 013's eventual oracle-driven gate** consumes the kind of
  observation reports this spec ships. Each release the per-format
  reports get richer (per-part diffs in 0.6.8, larger corpora in
  0.6.9, repair categorization in a later release).

## Normative References

- Spec 022 (first baselines, 0.6.6) — the methodology this extends.
- Spec 023 (auto-dismiss, 0.6.7) — the precondition for larger walks.
- Spec 024 (shared package_diff, 0.6.8) — the per-part-diff
  enrichment.
- `tools/oracle/baselines/README.md` "2026-04-29 second run" section
  — operational details, aggregate table, follow-up patterns.

## Approach

### Phase 1 — second baseline run + corpus caveat (this release, 0.6.9)

1. **Stage corpora** from `../tokenmoulds/{reports/visual, generated,
   scratch}` plus a small fresh-from-API run via
   `examples/build_acme_*.py` (which produces `.dotx` and `.xltx`
   templates — useful evidence for the template-roundtrip edge case
   documented in Spec 021).
2. **Run each oracle** via the new meta-CLI
   (`python -m openxml_audit.oracle <engine> <corpus>`) — the first
   release-grade exercise of the unified entrypoint shipped in
   0.6.8.
3. **Commit observation JSONs** under
   `tools/oracle/baselines/<format>/2026-04-29-larger.json`.
4. **Document the corpus-quality caveat** in
   `tools/oracle/baselines/README.md` so the elevated repair-dialog
   rate isn't read as a load-bearing claim about TokenMoulds'
   actual emitter output.
5. **Catalog the follow-up patterns** the run exposed (Word's
   "Open and Repair" button label edge case, PowerPoint's
   "[Repaired]" secondary modal, Excel template behavior) in the
   README's "Operational takeaways for follow-up patterns" section.

### Phase 2 — clean re-baseline against fresh TokenMoulds API output (later)

The right corpus is the output of TokenMoulds' actual emitters
driven by its Python API — `.docx` / `.xlsx` / `.pptx` / `.odt`
files produced as documents (not templates), with current
brand/font/locale defaults, written to a fresh dir per run. This
spec defers that work because:

- TokenMoulds' `examples/build_acme_*.py` produce templates
  (`.dotx`, `.xltx`, `.ott`), not concrete documents.
- The `tokenmoulds` CLI surfaces a `--generate-templates` /
  `--generate-odf` flag pair but errors on font-metrics-cache
  resolution under default install — needs more setup than this
  release accommodates.
- A "drive TokenMoulds programmatically to emit documents" workflow
  needs a small helper in `tools/oracle/` that's its own work.

Phase 2 builds that helper, runs it, gets a clean N=20+ per-format
baseline, and supersedes the 2026-04-29-larger reports.

## Acceptance Criteria (Phase 1 / 0.6.9)

1. `tools/oracle/baselines/{word,odf,pptx,xlsx}/2026-04-29-larger.json`
   exist and conform to the `RoundtripObservation` schema with the
   `added_parts` / `removed_parts` / `diff_dir` fields shipped in
   0.6.8 (or omitted gracefully when the diff didn't run).
2. The 2026-04-29-larger run was driven via
   `python -m openxml_audit.oracle <engine>` — first release-grade
   use of the meta-CLI from Spec 024.
3. `tools/oracle/baselines/README.md` documents the run, aggregate
   results, corpus-quality caveat, and the operational takeaways.
4. CHANGELOG updated. Spec 025 committed.
5. No code regressions; existing tests pass.

## Out of Scope (Phase 1)

- Phase 2's clean TokenMoulds-API-driven corpus (deferred).
- Pattern-list expansions for the dialog wordings observed in this
  run (catalogued, not implemented — they'd need cross-version
  Office testing to validate).
- The "[Repaired]" badge detection in PowerPoint titles
  (catalogued).
- Large-N corpus walks across non-TokenMoulds origins (e.g., LibreOffice
  qa-fixtures, real customer files). Spec 022's "Phase 2 corpus
  expansion" subsumes that.

## Risks

- **The headline numbers (preserved/repaired/open-failed counts)
  could be over-read** as load-bearing claims about Office app
  behavior on TokenMoulds output. Mitigation: explicit corpus
  caveat in the README.
- **The 4× larger walk takes substantially longer in wall-clock
  time** than the 0.6.6 walk. Most of that is per-file Office app
  open + close ceremony (~30-60s per file even on the happy path).
  Mitigation: don't run this on a dev machine you need responsive.
- **Repeated runs in the same session can leave Office apps in a
  fragile state** (we observed this with Excel after `pkill -9`,
  and again with PowerPoint after dialogs queued from a previous
  run). Mitigation: between corpus-walk sessions, force-quit the
  apps via AppleScript `quit saving no` rather than `pkill`.
