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
