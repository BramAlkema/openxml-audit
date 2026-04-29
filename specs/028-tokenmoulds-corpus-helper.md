# Spec: TokenMoulds-API Corpus Helper + Clean Re-baseline

## Status

Proposed (April 29, 2026). Phase 2 of Spec 026's roadmap to 0.8.0.
Resolves the corpus-curation gap flagged in 0.6.9 by *generating*
the corpus rather than *curating* it from existing TokenMoulds
filesystem artifacts. Replaces the 0.6.9 noisy baselines with a
clean re-run.

## Problem

The 0.6.9 (Spec 025) larger baseline assembled its corpus by globbing
`../tokenmoulds/{reports/visual,generated/word-demo,scratch/odf,...}`.
That corpus mixed release-quality TokenMoulds emitter output with
troubleshooting-session leftovers from the TokenMoulds project's own
development. The user flagged this mid-run; the 0.6.9 README
documented it as a known limitation. The headline finding from that
run — "TokenMoulds output triggers repair dialogs" — was loaded with
that caveat: maybe true of the leftover files, maybe not true of
TokenMoulds' actual emitter output.

This spec closes the gap by inverting the corpus-assembly
strategy: instead of picking files from disk, drive TokenMoulds'
emitter API directly and post-process the bytes into documents
the oracles can roundtrip end-to-end.

## Why This Matters

- **Provenance.** Every file in the new corpus is the byte-output
  of `tokenmoulds.emitters.<format>.<Emitter>(ir).build_package()`
  with a content-type swap. Anyone running the same script with
  the same brand inputs produces the same bytes — reproducible,
  no `git log`-fishing required.
- **Defensibility of headline findings.** "TokenMoulds Word/PPTX
  output is canonical to its target app" is a load-bearing claim
  with this corpus. Without the curation fix, that claim was
  contaminated.
- **Foundation for 0.7.3+ work.** Repair categorization (Phase 3
  of Spec 026's roadmap) and the eventual Spec 013 oracle gate
  both depend on a stable, reproducible corpus.

## Normative References

- `specs/026-self-parity-sovereign-gate-roadmap.md` — the umbrella
  this Phase 2 implements.
- `specs/025-larger-corpus-baselines.md` — Phase 1 of Spec 026's
  data-collection track; documented the corpus-curation gap.
- `tools/oracle/baselines/README.md` — extended in this release
  with the v0.7.2 clean run section, the Word/PPTX/Excel/ODF
  per-format findings, and the operational follow-ups
  (Excel external-source dialog handler).
- TokenMoulds' Python API:
  `tokenmoulds.dtcg.generator.TokenGenerator`,
  `tokenmoulds.ir.builder.build_document_ir`,
  `tokenmoulds.emitters.{word,excel,powerpoint}.<*Emitter>`,
  `tokenmoulds.emitters.odf.{Writer,Calc,Impress}TemplateEmitter`.

## Approach

### Phase 1 — corpus helper + re-baselines (this release, 0.7.2)

1. **`tools/oracle/build_corpus.py`** — driver that:
   - Constructs a minimal `DocumentIR` via `TokenGenerator` +
     `build_document_ir` (matching `examples/build_acme_*.py`
     and `tests/unit/emitters/test_powerpoint_emitter.py`).
   - Calls each emitter's `build_package()` for bytes.
   - Post-processes each one with `template_to_document(bytes,
     target=<format>)` — flips the content-type
     (`[Content_Types].xml` for OOXML; `mimetype` + manifest for
     ODF) from template to document variant. Lossless: the
     underlying XML is unchanged, only the format-identification
     bytes shift.
   - Writes 12 files (2 brand variants × 6 formats) under
     `data/corpus/tokenmoulds_v0.7.2/<format>/<brand>.<ext>`.

2. **Re-run all four oracles** on the new corpus via
   `python -m openxml_audit.oracle <engine>`. Commit observation
   JSONs under `tools/oracle/baselines/<format>/v0.7.2-clean.json`.

3. **Update `tools/oracle/baselines/README.md`** with the v0.7.2
   clean run section: methodology, aggregate table,
   per-format findings, operational follow-ups, and a side-by-side
   comparison vs the 0.6.9 mixed baseline.

### Phase 2 — bigger corpora, more brand variations (later)

The current driver emits 12 files (2 brand inputs × 6 formats).
A future release expands the brand-input matrix (more
locale/font/color combinations, additional org_id values) to grow
the corpus into the dozens or hundreds. The bottleneck is roundtrip
wall-clock time on Office apps, not corpus generation.

### Note on self-parity baseline (Spec 026 caveat)

Spec 026's 0.7.2 step originally said "re-run the self-parity
baseline (from 0.7.1) on this refreshed corpus." On reflection,
that's wrong — the self-parity baseline's purpose is to detect
**validator-output regressions**, which requires a *stable*
corpus that doesn't change. The SDK seed corpus (pinned to v3.4.1)
is that. The TokenMoulds corpus is for **app-survival oracle**
work and changes as TokenMoulds evolves. Running self-parity
against a moving corpus would conflate "validator changed" with
"corpus changed."

The 0.7.1 baseline against the SDK seed corpus stays as the
self-parity reference. This spec's corpus is exclusively for
oracle baselines.

## Headline findings (0.7.2 run)

| Format | Result | Insight |
|---|---|---|
| Word    | 2/2 preserved (3s each) | TokenMoulds output is canonical to Word |
| PPTX    | 2/2 preserved (3s each) | Canonical to PowerPoint |
| Excel   | 2/2 repaired (10 parts changed each) | **TokenMoulds Excel output is NOT canonical to Excel** — Excel rewrites 10 parts on every save, regardless of external-link updates |
| ODF     | 6/6 repaired | soffice canonicalizes everything (expected) |

The Excel finding is mission-relevant: the validator could
potentially detect what makes TokenMoulds' Excel output non-canonical
so users avoid the issue. With the per-part diffs landed in
0.6.8, the artifacts to answer "what specifically does Excel
rewrite?" are now collectible (see follow-up #2 in the README).

## Acceptance Criteria (Phase 1 / 0.7.2)

1. `tools/oracle/build_corpus.py` exists and produces 12 files
   across 6 formats under `data/corpus/tokenmoulds_v0.7.2/`.
2. The four oracle baselines under
   `tools/oracle/baselines/<format>/v0.7.2-clean.json` exist
   and follow the `RoundtripObservation` shape from 0.6.8.
3. `tools/oracle/baselines/README.md` documents the run,
   findings, and operational takeaways with a side-by-side
   comparison to the 0.6.9 mixed-corpus baseline.
4. CHANGELOG updated. Spec 028 committed.
5. No code regressions; existing tests pass.

## Out of Scope (Phase 1)

- **Larger brand-input matrix** (more locales, more color/font
  combos). The 12-file baseline is enough to demonstrate the
  approach; expansion can happen incrementally.
- **External-source / "Don't Update" dialog handler.** Cataloged
  as an operational follow-up; needs a separate detection path
  from the repair-dialog flow because the dialog isn't about
  repair.
- **Investigation of the 10 changed parts on Excel files.** The
  finding is committed; the per-part text diffs are in the
  observation reports under `<work_dir>/compare/diffs/` when
  invoked with `--keep-artifacts`. A focused study slots into
  0.7.3 or later.
- **Self-parity baseline rebase** (deliberately deferred per the
  caveat above).
- **Corpus regeneration with different TokenMoulds revs** (e.g.
  to detect TokenMoulds emitter regressions). The current
  baseline is at "TokenMoulds main as of 2026-04-29".

## Risks

- **TokenMoulds API drift.** If TokenMoulds renames `TokenGenerator`,
  `build_document_ir`, or any emitter class, the corpus builder
  breaks. Mitigation: TokenMoulds is a sibling project on the same
  dev machine; updates to its API trigger a parallel update here.
- **Font metrics cache dependency.** The driver requires
  TokenMoulds' committed `data/fonts/cache/metrics.json` to exist;
  resolved relative to TokenMoulds' module path. If TokenMoulds
  is installed via pip-from-PyPI without that data file, the
  builder errors. Mitigation: a future release vendors a minimal
  metrics cache here.
- **Content-type swap may break in TokenMoulds revs that change
  how templates self-identify.** Mitigation: the swap is a
  single bytes-replace per format; trivially adjustable when
  TokenMoulds changes its template content type strings.
- **The 12-file corpus is small.** A regression could pass
  unnoticed if it only manifests on specific brand inputs.
  Mitigation: this is Phase 1; Phase 2 expands the matrix.
