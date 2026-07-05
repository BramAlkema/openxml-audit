# Roundtrip Oracle Baselines

Per-format observation snapshots from running the four roundtrip oracles
(`tools/oracle/{word,odf,pptx,xlsx}_repair_oracle.py` / `word_repair_corpus.py`)
against curated corpora. Each subdirectory holds dated JSON reports
matching the schema in `RoundtripObservation`.

## Layout

```
tools/oracle/baselines/
  README.md                  ← this file
  word/<YYYY-MM-DD>.json
  odf/<YYYY-MM-DD>.json
  pptx/<YYYY-MM-DD>.json
  xlsx/<YYYY-MM-DD>.json
  word_<constraint>_pairwise.json   ← scenario-matrix oracle output
  (e.g. word_trpr_pairwise.json, word_tblpr_pairwise.json, ...)
```

The dated files come from corpus-driven oracle runs (`*_repair_corpus.py`
or `*_repair_oracle.py` taking files/dirs as input). The
`word_<constraint>_pairwise.json` files come from the matrix-driven Word
oracle (`word_repair_oracle.py <constraint>`) which generates synthetic
scenarios and runs them through Word — a different shape of oracle, same
output schema.

## Reproducing

Each baseline JSON records the inputs, outcomes, and per-file durations.
To reproduce:

```bash
python tools/oracle/odf_repair_oracle.py  /path/to/odf/corpus  --output ...
python tools/oracle/word_repair_corpus.py /path/to/docx/corpus --output ...
python tools/oracle/pptx_repair_oracle.py /path/to/pptx/corpus --output ...
python tools/oracle/xlsx_repair_oracle.py /path/to/xlsx/corpus --output ...
```

Permissions setup is required for the three Microsoft Office oracles —
see `docs/oracle_permissions.md`. The ODF oracle uses headless soffice
and needs no special grants.

## 2026-04-29 baseline run — first across all four formats

| Format | Files | Preserved | Repaired | Crashes | Repair dialog seen |
|---|---|---|---|---|---|
| ODF | 5 | 0 | 5 | 0 | n/a (soffice silent) |
| Word | 1 | 1 | 0 | 0 | 1 |
| PPTX | 2 | 2 | 0 | 0 | 0 |
| Excel | 1 | 1 | 0 | 0 | 1 |
| **Total** | **9** | **4** | **5** | **0** | **2** |

### Findings

**ODF (5 files, all `repaired`).** soffice canonicalizes formatting on
every save — every file's canonical parts (`content.xml`, `styles.xml`,
`meta.xml`, `settings.xml`) come back with different bytes than they
went in. This is expected; soffice's repair surface is silent (no
modal alerts in headless mode), and the diff just records "the bytes
changed." Phase 2 work in Spec 019 will categorize the diffs
(cosmetic-vs-substantive) once we have an XML-level differ.

**Word (`mcp-word.docx`, `preserved`, dialog seen).** Word displays
its repair dialog on every open of this file ("Word was unable to read
this document. It may be corrupt"). The oracle clicks through with
"Open and Repair" (a dialog wording / button label added during this
baseline run), Word repairs, and the post-repair canonical XML
matches the input byte-for-byte. **Real signal: Word silently repairs
this file every open without changing what's on disk** — exactly the
class of issue the validator is built to detect.

The repair-dialog wording added to `word_window.REPAIR_DIALOG_PATTERNS`
in this release (`unable to read this document`, `may be corrupt`,
`open and repair`, `text recovery converter`) and the `Open and Repair`
button label added to `word_roundtrip._try_dismiss_repair_dialog`'s
alt list extend the Word oracle's coverage.

Earlier in this run the corpus also included `tokenmoulds/scratch/demos/*`
.docx files — those triggered the same dialog repeatedly and were
dropped because they're scratch/dev-broken artifacts, not realistic
corpus material.

**PowerPoint (2 files, all `preserved`).** PowerPoint roundtrip is the
cleanest of the four — both `winsemius.pptx` and `mcp-swot.pptx` open,
close-with-save, and come back byte-identical on the canonical parts
the oracle fingerprints. No repair dialogs surfaced. PowerPoint's
canonical-form discipline appears stricter than Excel's or Word's
on TokenMoulds output.

**Excel (`winsemius.tokens.xlsx`, `preserved`, dialog seen).** Excel
displays the "We found a problem with some content" repair dialog
on this file. The oracle's pattern list catches it; user accepted the
recovery (auto-dismiss is Phase 2 of Spec 021). Excel's followup
dialog ("Excel was able to open the file by repairing or removing the
unreadable content") was captured in `repair_dialog_text`. After
recovery, Excel close-with-save produced byte-identical canonical
parts — same as Word's behavior on `mcp-word.docx`. **Strong signal
for the validator's mission: TokenMoulds-emitted Excel files trigger
Excel's repair flow.**

### Operational issues fixed during this run

- `pptx.osa.list_open_presentation_names` and
  `xlsx.osa.list_open_workbook_names` were using the
  `repeat with X in COLLECTION / name of X` AppleScript idiom, which
  hangs the AppleScript engine indefinitely on Office for Mac M365 16.x
  (a known issue documented in `tools/oracle/word_window.py`'s comment
  about "infinite loops inside nested tells"). Both now use
  `name of every X` — the same idiom Word's oracle has used since
  Spec 011.
- `word_window.REPAIR_DIALOG_PATTERNS` extended with the "file is
  corrupt" dialog wording. Without this addition the oracle would
  hang on Word's hard-error dialog.
- `word_roundtrip._try_dismiss_repair_dialog`'s alt-button list
  extended with `("Open and Repair", "Cancel")`.

## 2026-04-29 second run — larger corpus, partial completion

A second run on the same date with an expanded corpus, queued under
`{word,odf,pptx,xlsx}/2026-04-29-larger.json`. The motivation was Phase
2 of Spec 022 (more files per format) made tractable by Spec 023's
auto-dismiss in 0.6.7. The 0.6.8 shared per-part differ now annotates
the observations with `added_parts` / `removed_parts` / `diff_dir` —
real per-part signal where 0.6.6 had hash-only.

### Aggregate

| Format | Files | Preserved | Repaired | Open-failed | Repair dialog seen |
|---|---|---|---|---|---|
| ODF | 7 | 0 | 1 | 6 | n/a |
| Word | 4 | 2 | 0 | 2 | 1 |
| PPTX | 5 | 3 | 0 | 2 | 2 |
| XLSX | not re-run | — | — | — | — |

### What the data says

- **3/4 PPTX files came back `preserved`** — the auto-dismiss path (0.6.7) caught the repair dialog on 1 of 2 problematic files cleanly.
- **2/4 Word files came back `preserved`** — `table-test.docx` triggered a repair dialog and the auto-dismiss successfully clicked through.
- **The remaining open-failed cases all hit dialog variants the pattern lists don't yet cover**, plus PowerPoint's secondary "[Repaired]" info modal which has button labels (`View` / `Delete` / X) outside our `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`. The post-close-with-save Escape pass added in 0.6.7 catches this on the next file but not on the one that's currently mid-roundtrip.

### Corpus-quality caveat (important)

The 2026-04-29-larger corpus was assembled by globbing
`../tokenmoulds/{reports/visual,generated/word-demo,scratch/odf,...}`
without distinguishing **finished, release-quality TokenMoulds emitter
output** from **leftovers from troubleshooting sessions during
TokenMoulds development**. The user flagged this midway through the
Word run: several of the high-friction files (the `Word was unable to
read this document` cases on `mcp-word.docx` and `demo-workflow.docx`)
likely come from the latter category, not from TokenMoulds' current
emitters.

That has two consequences:

1. **The headline "TokenMoulds output triggers repair dialogs" finding
   from the 2026-04-29 (first) baseline should be read with this caveat
   too** — `winsemius.tokens.xlsx` and `mcp-word.docx` were corpus
   choices made by the same indiscriminate process. The fact that
   they trigger dialogs is real; whether that says anything about
   TokenMoulds' production-quality emitter output is unclear from
   this corpus.
2. **The operational findings (dialog wordings, button labels,
   auto-dismiss gaps) are real regardless** — those are properties
   of how Word/Excel/PowerPoint behave on imperfect input, and the
   oracle's job is to capture them faithfully.

A clean re-baseline against TokenMoulds' actual production output
(driven through its Python API or CLI to produce concrete
`.docx`/`.xlsx`/`.pptx`/`.odt` rather than templates) is deferred
to a future release. For now the per-part diff machinery, the
auto-dismiss, and the meta-CLI all work; the corpus is the missing
piece, not the tooling.

### Operational takeaways for follow-up patterns

These dialog variants appeared during the run and are candidates for
adding to the pattern lists:

- **Word "unable to read" hard-error**: button "Open and Repair"
  was added in 0.6.6 but didn't always click successfully across
  the run. Worth re-checking — possibly the System Events button
  label has different capitalization or a non-breaking space in
  some Word builds.
- **PowerPoint "[Repaired] and removed it. View / Delete"
  post-repair info**: dismiss path's `dismiss_any_leftover_modal`
  Escape pass should clear this at end-of-roundtrip, but if the
  modal appears mid-roundtrip the auto-dismiss tries OK / Yes /
  Repair / Recover / Open and finds none. Could add `View` (which
  opens the log but doesn't lose the file) or `Close` to the
  accept list, OR detect the "[Repaired]" badge in the title and
  Escape directly.
- **Excel post-`pkill` recovery sheet**: same as the 2026-04-29
  first run; benign.

These are grist for 0.7.x and beyond, not blockers for the 0.6.9 ship.

## v0.7.2 clean run — TokenMoulds-API corpus

A re-baseline against a corpus that's *generated* rather than
*curated*: `tools/oracle/build_corpus.py` drives TokenMoulds'
emitter API directly to produce 12 documents (2 per format ×
6 formats), then post-processes each one to flip its content
type / mimetype from template (`.dotx` / `.xltx` / `.potx` /
`.ott` / `.ots` / `.otp`) to document
(`.docx` / `.xlsx` / `.pptx` / `.odt` / `.ods` / `.odp`).

The corpus lives under
`data/corpus/tokenmoulds_v0.7.2/{word,excel,pptx,odf}/` and is
byte-reproducible from the same brand inputs run-to-run. The user-
flagged corpus-curation problem from the 0.6.9 run is resolved at
the source: every file in this corpus is provably "what TokenMoulds
currently emits," nothing more, nothing less.

### Aggregate (12 files across 4 oracles)

| Format | Files | Preserved | Repaired | Open-failed | Repair dialog seen |
|---|---|---|---|---|---|
| Word    | 2 | **2** | 0 | 0 | 0 |
| PPTX    | 2 | **2** | 0 | 0 | 0 |
| Excel   | 2 | 0     | **2** (10 changed parts each) | 0 | 0 |
| ODF     | 6 | 0     | 6 (soffice canonicalizes — expected) | 0 | n/a |

### Findings

- **Word and PowerPoint accept TokenMoulds output as canonical.**
  2/2 files in each format roundtripped in 2-3 seconds with zero
  byte changes on the parts the oracle fingerprints. The 0.6.9
  baseline's open-failed Word and PPTX cases were the *corpus*,
  not the emitter — confirmed.

- **Excel repairs every TokenMoulds-emitted workbook.** Both
  `.xlsx` files came back with 10 changed parts per file. No
  repair dialog (Excel chose to silently rewrite rather than
  prompt). This is the headline mission-relevant signal of the
  whole run: TokenMoulds' Excel emitter produces workbooks that
  are valid by openxml-audit's schema/semantic rules but get
  rewritten on every Excel save. The 10 changed parts (committed
  under `…/v0.7.2-clean.json`) include the canonical XL parts —
  which means Excel disagrees with TokenMoulds about something
  fundamental in the package. Worth a focused investigation in a
  later release.

- **External-source dialog observed for the first time.** Excel
  showed a "This workbook contains links to one or more external
  sources that could be unsafe. ... Don't Update / Update" modal
  on `acme-us.xlsx` (33.6s including the manual click — the user
  clicked Update). The globex run cleared in 5.1s without the
  prompt — the dialog is per-file, dependent on what TokenMoulds
  emitted into that specific workbook. Buttons aren't in the
  current `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`; needs a separate
  detection path because this isn't a repair flow.

  Note that BOTH files came back with 10 changed parts, despite
  only acme-us getting the dialog. So the 10-parts diff is mostly
  Excel's canonicalization on save (per-file deterministic), not
  the external-link update — that update only ran for acme-us
  and produced no measurable extra delta on the parts we
  fingerprint. **The "Excel rewrites every TokenMoulds-emitted
  workbook" finding holds independent of the link-update path.**

- **soffice flakiness from earlier today resolved.** The 0.6.9
  ODF run had 6/7 open-failed; this run had 6/6 succeed. Same
  harness, different LibreOffice state. Reinforces "between
  corpus walks, force-quit Office apps via AppleScript rather
  than `pkill -9`" from the previous README section.

### Operational takeaways for follow-up

1. **External-source / "Don't Update" dialog handler.** Excel's
   "external sources" modal needs detection separate from the
   repair-dialog path. Default action should be "Don't Update"
   (preserve the document state for the roundtrip).
2. **Excel canonical-form gap.** The 10-changed-parts finding
   warrants a focused study: which parts is Excel rewriting?
   Cosmetic XML reflow, or substantive structural changes? With
   the per-part diffs landed in 0.6.8, the artifacts to answer
   this are now collectible.
3. **PowerPoint "[Repaired]" secondary modal** (from 0.6.9) was
   not exercised on this corpus because no file triggered the
   primary repair flow. Still on the follow-up list.

### Comparing the v0.7.2 clean baseline vs the 0.6.9 mixed corpus

| | 0.6.9 | 0.7.2 (clean) |
|---|---|---|
| Word success rate | 2/4 (50%) | 2/2 (100%) |
| PPTX success rate | 3/5 (60%) | 2/2 (100%) |
| Open-failed cases | 4 of 9 OOXML | 0 of 6 OOXML |
| Headline narrative | "TokenMoulds output triggers repair dialogs" | "TokenMoulds Word/PPTX clean; Excel rewrites every save" |

The corpus-curation gap, fixed.

### Excel post-kill recovery dialog

Force-killing Excel mid-roundtrip (e.g. `pkill -9 "Microsoft Excel"`
during oracle debugging) puts Excel into a "Microsoft Excel will
attempt to recover your work" state on the next launch. This is a
macOS Excel feature, not an oracle issue — Excel writes auto-recovery
files under `~/Library/Containers/com.microsoft.Excel/Data/Library/
Application Support/Microsoft/Office/UnsavedFiles/` whenever it
unexpectedly terminates and prompts the user to recover them on next
open.

The recovery prompt is benign for oracle work — there's no real
unsaved data because the oracle's close-with-save already persists
to the staged path. Click "Don't Recover" (or "Recover", same effect
for our purposes) and proceed. To avoid the prompt entirely in
debugging sessions, prefer `osascript -e 'tell application "Microsoft
Excel" to quit saving no'` over `pkill -9`.

### Excel template (`.xltx`) limitation

Excel template files don't roundtrip via the oracle's open→save
pattern. Excel opens a `.xltx` by *creating a new untitled workbook*
based on the template, not by opening the template itself. The
oracle's `is_workbook_open(staged)` polls for the staged filename and
times out because Excel never registers the staged file as open.

This is a documented Excel behavior, not an oracle bug. To roundtrip
templates, the oracle would need a different code path: open template,
detect the new untitled workbook by `name of active workbook`, save it
as a new `.xltx` using AppleScript's `save active workbook as` with
the `xlt` format constant. Deferred to Spec 021 Phase 2.

### What this baseline does NOT prove

- That the *oracle classification* is right — `preserved` is a
  byte-level claim about the parts we fingerprint, not a guarantee
  that the file is semantically identical. Word's repair dialog
  showing AND the canonical parts being byte-identical post-repair
  is suspicious; in 0.6.7's per-part text differ we'll see if the
  parts we DON'T fingerprint (docProps, settings, etc.) differ.
- That TokenMoulds output is broken — Word and Excel showing repair
  dialogs is real signal, but it's not yet diagnosis. The next step
  is feeding these files through this validator to see if our
  schema/semantic checks predict the repairs.
- That the corpus is representative — 9 files across 4 formats is a
  starter, not a comprehensive walk.

## 2026-07-05 — first LibreOffice OOXML roundtrip baseline (Spec 036)

`lo_ooxml/2026-07-05.json` — first run of the CI-runnable oracle
(`tools/oracle/lo_ooxml_repair_oracle.py`, LibreOffice 26.2.4.2 on
macOS, headless soffice). Corpus: TokenMoulds v0.7.2 OOXML files (6)
plus the three Spec 034 reference documents built at
`--minimum-tier loadable` from this commit.

| Outcome | Count |
|---|---|
| Opened + roundtripped ("repaired") | 9/9 |
| Crash / timeout / open-failed | 0 |
| Byte-preserved | 0 (expected — LO always rewrites foreign formats) |

### Findings

- **LibreOffice opens everything in this corpus**, including the
  generated reference documents — no repair prompts (headless is
  silent), no crashes, 2-5s per file.
- **Feature survival on `reference.pptx`** (per-feature structural
  probes from `openxml_audit.reference.feature_probes`):
  entrance fade, entrance wipe, click end-condition, time-offset
  end-condition, and non-root `restart` all survive the roundtrip
  structurally. **`repeatDur` does not survive**: LibreOffice keeps
  `repeatCount` and silently drops the `repeatDur` cap (verified by
  hand in the converted slide — the only remaining "repeatDur" string
  is the probe slide's label text). First live-app feature-survival
  result produced by the Spec 034 → Spec 036 ladder cycle.
- Part-level changes are universal (2-9 changed canonical parts per
  file) and expected; categorizing cosmetic vs substantive rewrites
  remains Spec 019 Phase 2 territory.

Scope note: these observations are about **LibreOffice as a target
app**. They say nothing about PowerPoint/Word/Excel behavior and do
not feed capability-registry tier promotion.
