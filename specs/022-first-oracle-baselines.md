# Spec: First Oracle Baselines Across All Four Formats

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.6 — first
baselines committed under `tools/oracle/baselines/{word,odf,pptx,xlsx}/`,
operational fixes that surfaced during the run, README documenting
methodology + findings.

## Problem

0.6.1 → 0.6.5 built four roundtrip oracles (Word, ODF, PowerPoint,
Excel). Up through 0.6.5 they were *infrastructure* — tools that
existed but had never produced data. The 0.7.0 narrative would have
been "we built four tools," which is hollow without evidence the
tools actually observe anything.

This release runs each oracle on a real (TokenMoulds-emitted) corpus,
commits the JSON observation reports, and documents what we learned.
The oracle ladder transitions from "shipped" to "demonstrated."

## Why This Matters

- **Evidence over infrastructure.** A roundtrip oracle that has never
  been run on real input proves nothing. A small set of committed
  baselines proves the tooling works on actual files and produces
  meaningful signal.
- **Operational shakedown.** First-use baselines surface bugs that
  unit tests with synthetic input miss — and they did:
  AppleScript-idiom hangs, missing repair-dialog patterns, button-label
  gaps, Excel template (`.xltx`) edge cases. All fixed in this release.
- **Spec 013 readiness.** The eventual self-parity gate (Spec 013) needs
  to know what oracle observations look like in practice. Committed
  baselines are the reference.
- **Marketing/positioning material.** "We built 4 oracles AND ran them
  on N files; here's what we found" is shippable narrative for 0.7.0.
  "We built 4 oracles" alone is not.

## Normative References

- `specs/011-word-roundtrip-oracle.md` (Word oracle, since 0.5.0)
- `specs/019-odf-roundtrip-oracle.md` (ODF, 0.6.3)
- `specs/020-pptx-roundtrip-oracle.md` (PowerPoint, 0.6.4)
- `specs/021-xlsx-roundtrip-oracle.md` (Excel, 0.6.5)
- `tools/oracle/baselines/README.md` — methodology, results table,
  per-format findings, operational fixes from this run
- `docs/oracle_permissions.md` — macOS setup gate

## Approach

### Phase 1 — collect first baselines (this release, 0.6.6)

1. **Stage corpora** from TokenMoulds-emitted output:
   - ODF: 5 files (`.odt`, `.odp` mix from `tokenmoulds/generated/`).
   - Word: `mcp-word.docx` (the `tokenmoulds/scratch/demos/` set was
     attempted first and rejected — those are dev-broken artifacts
     that hang Word's repair dialog without ever opening).
   - PowerPoint: `winsemius.pptx`, `mcp-swot.pptx`.
   - Excel: `winsemius.tokens.xlsx` (the `.xltx` template files were
     attempted first; Excel doesn't open them as workbooks — see
     limitation note in `tools/oracle/baselines/README.md`).
2. **Add `tools/oracle/word_repair_corpus.py`** — sibling of the
   ODF/PPTX/XLSX corpus walkers. The existing
   `word_repair_oracle.py` is a scenario-matrix runner (Spec 010 / 011
   phase work); this new file walks arbitrary `.docx` corpora using
   `word_roundtrip.roundtrip()` and emits the same
   `RoundtripObservation` shape as the other three oracles.
3. **Run each oracle** with timeout 90s, output to
   `tools/oracle/baselines/<format>/<YYYY-MM-DD>.json`.
4. **Fix operational issues found during the run** (see Findings).
5. **Document everything** in `tools/oracle/baselines/README.md`.

### Phase 2 — corpus expansion + structural categorization (later)

- Larger TokenMoulds-driven corpora (50+ files per format).
- Per-part text diffs (`pptx.lab.compare_pptx_packages` extracted to a
  shared `openxml_audit.parts.diff`) replace hash-only outcomes.
- Repair-vs-cosmetic categorization on the diffs.
- Auto-dismiss for PPTX and XLSX repair dialogs (Word has it; PPTX/XLSX
  Phase 2 of their respective specs).
- LibreOffice qa import as secondary ODF corpus.

## Findings (Phase 1)

See `tools/oracle/baselines/README.md` for the full table and
per-format breakdown. High points:

- **All 4 oracles work on real files.** 9 files across 4 formats,
  zero crashes, zero timeouts, zero hard failures.
- **TokenMoulds-emitted Word and Excel files trigger repair dialogs
  in their target apps**, even though the post-repair canonical XML
  is byte-identical to the input. This is the validator's exact
  moat: predict these dialogs without needing to run Word or Excel.
- **PowerPoint roundtrip is the cleanest** — 2/2 files survived
  unchanged, no repair dialogs.
- **soffice canonicalizes ODF on save** — every file marks
  `repaired`, but that's expected since soffice rewrites all parts.

### Operational fixes that shipped during this run

1. `pptx.osa.list_open_presentation_names` and
   `xlsx.osa.list_open_workbook_names` switched from
   `repeat with X in COLLECTION / name of X` to `name of every X`
   — the broken idiom hangs Office for Mac M365's AppleScript
   engine indefinitely (Word's oracle already knew this; PPTX/XLSX
   inherited the wrong pattern).
2. `word_window.REPAIR_DIALOG_PATTERNS` extended with the
   "file is corrupt → Open and Repair?" dialog wording. Without
   this, the oracle hung on the hard-error dialog.
3. `word_roundtrip._try_dismiss_repair_dialog` alt-button list
   extended with `("Open and Repair", "Cancel")`.

## Acceptance Criteria (Phase 1 / 0.6.6)

1. `tools/oracle/baselines/<format>/<YYYY-MM-DD>.json` files exist
   for all four formats (Word, ODF, PPTX, XLSX).
2. `tools/oracle/baselines/README.md` documents methodology, results,
   and findings.
3. `tools/oracle/word_repair_corpus.py` exists and walks arbitrary
   `.docx` corpora with the same observation contract as ODF/PPTX/XLSX.
4. Operational fixes (`name of every X` idiom, repair-dialog patterns,
   button labels) are committed and don't regress existing tests.
5. CHANGELOG updated. Spec 022 committed.

## Out of Scope (Phase 1)

- Larger corpora (Phase 2).
- Per-part text diff (Phase 2; needs `pptx.lab` differ extraction).
- Repair categorization (Phase 2).
- Auto-dismiss for PPTX/XLSX (Phase 2 of Specs 020/021).
- `.xltx` / `.dotx` / `.potx` template-roundtrip path (different code
  flow; deferred).
- Wiring the oracles into CI (Spec 013's job).

## Risks

- **One file per format isn't representative.** Mitigation: this is
  Phase 1 — the goal is "the tooling works on real input," not
  comprehensive coverage. Phase 2 expands.
- **Hash-based diff collapses cosmetic and substantive repairs.** A
  Word file showing the repair dialog AND coming back byte-identical
  on the parts we fingerprint is suspicious; the docProps or settings
  parts (which we don't fingerprint) may differ. Phase 2's per-part
  text diff narrows this gap.
- **TokenMoulds is one author family.** Files emitted by Microsoft
  Office, Pages, Google Docs, etc., have different corner cases.
  Phase 2 expands corpus origins.
- **macOS automation surface remains fragile.** Each release run
  surfaces new dialog wordings, button labels, sandbox quirks. The
  pattern lists and alt-button lists are growing organically as
  evidence accumulates — that's the project working as designed,
  not a flaw.
