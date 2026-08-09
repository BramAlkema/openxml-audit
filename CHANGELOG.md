# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.7.8] - 2026-08-09

Euro-Office conversion evidence and a packaging repair for the public oracle
dispatcher. PyPI installs can now run every dispatcher engine outside a source
checkout, including the new authenticated Euro-Office conversion oracle.

### Added
- **Euro-Office conversion oracle (Spec 038)** — a dependency-free
  `openxml_audit.eurooffice` client and `openxml-audit-oracle eurooffice`
  engine for the current Document Server `/converter` contract. The format
  model is pinned to Nextcloud connector 11.0.1 and document-formats 3.2.0:
  native OOXML editing, lossy ODT/ODS/ODP-to-OOXML editing, and ODG
  view/conversion only. JWT secrets are environment-only, returned OOXML is
  validated, and same-format conversions receive canonical per-part diffs.
  Reports state their limit explicitly: this proves conversion, not browser
  editing, save callbacks, or coauthoring. The first six-file live baseline
  records two validator-clean native conversions, three returned packages with
  validator findings (including unchanged source findings for PPTX), and one
  ODP conversion failure; it also exposes the live 9.3.1.37 versus upstream
  9.3.3 deployment-version gap.

### Fixed
- **Packaged oracle dispatcher** — wheels now include the `tools.oracle`
  implementation package used by `openxml-audit-oracle`. PyPI installs no
  longer fail with `ModuleNotFoundError: No module named 'tools'` when an
  engine is selected outside a repository checkout.

## [0.7.7] - 2026-07-20

A reported bug that turned out to be the visible corner of a larger
one: rules the validator loads, converts and reports as covered, but
which can never match anything.

### Added
- **Cross-part `wp:docPr` id uniqueness (issue #6)** — Word renumbers a
  drawing whose id collides with one in another part of the same
  document story. The SDK's schematron carries the rule but scopes it
  to a single part, so a collision straddling `document.xml` and a
  footer satisfies it. Scope is the main document plus the
  header/footer parts it references; the glossary is excluded as a
  separate id space. WARNING severity and `WORD_APP_COMPAT`: Word
  repairs the file rather than refusing it, and the verdict layer still
  reports `repair-or-rewrite-likely`.
- **Spec 037 — enforcement-fidelity invariants**, with
  `scripts/audit_constraint_enforcement.py`: 77 of 213 bridged
  `UNIQUE_ATTRIBUTE` constraints never fire, because attribute
  qualification is decided by a two-entry allowlist rather than the SDK
  data. Three existing guards are blind to it — coverage asserts
  *converted*, spec 009's invariants cover only the schema bridge, and
  parity is advisory over mostly-valid files. The spec argues against an
  overhaul (spec 033/ADR-004 already shelved one on measurement) in
  favour of one source of qualification truth plus an executability
  invariant.

### Fixed
- **Inherited SDK attributes are now resolved.** The SDK declares an
  attribute once, on the type that introduces it; nothing walked the
  `BaseClass` chain, so elements whose attributes are all inherited —
  `a:rPr`, `a:defRPr`, `a:pPr` among them — presented an empty
  attribute list. That disabled undeclared-attribute detection for 781
  element tags and left 1513 typed attributes unvalidated:
  `<a:rPr sz="NOT_A_NUMBER"/>` validated clean. Resolution is scoped per
  schema, because `ClassName` is not unique across namespaces (385
  collide) and a global index attaches SpreadsheetML attributes to
  DrawingML elements.
- **`w:start` no longer resolves to the border type.** It ties with
  `CT_Border` on every scoring component and the final tiebreak prefers
  whichever type declares more attributes, so `<w:start w:val="1"/>` in
  `numbering.xml` was checked against the page-border-art enumeration.
  Now resolved by parent, as the other overloaded Word tags already
  are. The mis-selection predated this release; it was invisible while
  `CT_Border` had no attributes to apply.
- **Security audit workflow** upgrades setuptools, so the weekly scan
  stops going red on ambient runner tooling unreachable from this
  project.

### Notes
- Findings over the corpus and fixtures are unchanged by this release:
  17 before, 17 after, none new and none lost.

## [0.7.6] - 2026-07-05

The reference-document arc: the canonical per-format reference
documents ADR-001 named as the payoff now exist as a generated,
tier-honest pipeline — plus the first CI-runnable live-app oracle and
an app-survival verdict as the CLI's headline answer.

### Added
- **Canonical reference documents (Spec 034)** — `openxml_audit.reference`:
  cross-format capability ledger, per-format builders (PPTX from
  committed oracle scaffold slides behind a generated index slide;
  DOCX/XLSX as canonical packages), provenance manifests with
  machine-readable exclusion reasons, mandatory self-validation, and
  byte-reproducible builds. CLI: `python -m openxml_audit.reference
  build|status`.
- **Generated entrance fade/wipe reference slides** — authored at build
  time from the oracle decks' own fragment builders
  (`reference/pptx_slides.py`); all six registered PPTX findings now
  have emitters.
- **App-survival verdict (Spec 035)** — `openxml_audit.verdict` maps
  findings to a per-app prediction (opens-clean / at-risk /
  repair-or-rewrite-likely / reject-likely) via the SourceClass
  taxonomy. Text CLI output leads with the verdict headline; JSON
  output gains a `verdict` block. App-compat findings drive the
  verdict at any severity (the warning-severity Spec 029
  shared-strings rule predicts the rewrite the v0.7.2 oracle baseline
  confirmed).
- **LibreOffice OOXML roundtrip oracle (Spec 036)** — the CI-runnable
  rung: `tools/oracle/lo_ooxml_repair_oracle.py` on the supervised
  soffice harness (now supporting docx/xlsx/pptx targets), with
  per-feature survival probes (`reference/feature_probes.py`) driven
  by the reference manifest. First baseline committed
  (`tools/oracle/baselines/lo_ooxml/2026-07-05.json`): 9/9 files open;
  fade, wipe, end-conditions, and restart survive LibreOffice Impress;
  **repeatDur is silently dropped** (repeatCount kept). Advisory
  workflow `.github/workflows/reference-oracle.yml` runs it weekly and
  on reference-path PRs.

### Fixed
- Committed `timing_oracle` scaffold slides 2-3 placed
  `<p:endCondLst>` after `<p:childTnLst>`, violating
  CT_TLCommonTimeNodeData order — caught by the reference builder's
  self-validation; fixed in the deck builder and the committed XML.
- Reference manifest locations are span-aware ("slides 7-8") for
  multi-slide features.
- `__version__` drift (`0.7.4` in `openxml_audit/__init__.py` vs
  `0.7.5` released) corrected.

### Security
- lxml floor raised to >=6.1.0 (CVE-2026-41066 / PYSEC-2026-87: XXE
  local-file read via default parser; reachable here because the
  validator parses untrusted OOXML/ODF input).
- Weekly advisory `pip-audit` workflow
  (`.github/workflows/security-audit.yml`); `pip-audit` in dev extras.
- Multi-candidate constraint-resolution cache key now includes the
  parent namespace, preventing cross-namespace context aliasing.

## [0.7.5] - 2026-05-03

Added Google Suite to the mix.

### Added
- Google Workspace roundtrip oracle (Spec 031) — fifth engine in the
  oracle dispatcher alongside Word, ODF, PowerPoint, Excel.
  - **`src/openxml_audit/gsuite/`** — new package with `GSuiteClient`
    wrapping four Drive API primitives (`upload`,
    `convert_to_native`, `export_to_ooxml`, `delete`) plus mime-type
    constants for OOXML and native Google formats.
  - **`tools/oracle/gsuite_roundtrip.py`** — orchestrator that
    uploads an input .pptx, copies it with conversion to native
    Google Slides, exports the Slides file back to .pptx, hands
    original + roundtripped copy to `pptx.lab.compare_pptx_packages`
    for the per-part diff, then classifies the diff across a
    `LossClass` taxonomy.
  - **`LossClass` taxonomy** — naming policy: stay descriptive.
    `*_part_changed` / `*_part_removed` describe what's observable
    in the diff without claiming semantic loss; `*_loss` names are
    reserved for future buckets that verify actual loss (e.g.,
    `font_loss` once we detect "font removed AND referenced by a
    text run"). Phase 1 file-level buckets: `theme_part_changed`,
    `master_part_changed`, `style_part_removed`,
    `font_part_removed`, `slide_part_changed`, `metadata_churn`,
    `media_re_encoded`, `structural_normalization` (gain-side
    artifacts like Google adding notes masters), `defaults_inlined`
    (content-aware: source had empty `<p:spPr/>`/`<a:bodyPr/>`/etc.
    inheriting from layouts; head has them expanded to fully-resolved
    values — file-level only, doesn't claim semantic loss).
    Reserved verified-loss bucket: `content_changed`. Catch-all:
    `unmapped`. Multiple classes may fire per
    file. The taxonomy distinguishes loss (signal removed/degraded)
    from normalization (parts GSuite *added* that weren't in the
    source) — both are roundtrip artifacts but mean different things.
  - **Auth: domain-wide delegation (mandatory).** Service accounts
    have zero storage quota since Google's 2024 policy change, so
    `with_subject(...)` impersonation of a real Workspace user is
    required for any `files.create` call. Setup is a one-time
    Workspace Admin Console step (Security → API controls →
    Domain-wide Delegation) plus three env vars:
    `GSUITE_ORACLE_CREDS` (defaults to
    `~/.config/openxml-audit/google_service_account.json`),
    `GSUITE_ORACLE_SUBJECT`, `GSUITE_ORACLE_FOLDER_ID`.
  - **`pyproject.toml`** — new `gsuite` optional-dependencies group
    with `google-auth` only; install via `pip install -e ".[gsuite]"`.
    The Drive API HTTP layer is implemented with stdlib
    `urllib.request` (no `google-api-python-client`, no `requests`,
    no `urllib3`, no `protobuf`). google-auth handles JWT signing +
    token refresh; everything else is plain HTTP. Net dep cut: ~3 MB
    of transitive packages eliminated, ~80 lines of code added in
    `gsuite/client.py` to wrap Drive's REST endpoints + a 30-line
    `_UrllibAuthTransport` adapter that satisfies google-auth's
    transport interface during refresh().
  - **`oracle.__main__` dispatcher** — adds `gsuite` engine (alias
    `google`); usage: `python -m openxml_audit.oracle gsuite FILES...`.
  - **Empirical baseline.** A 32 KB single-slide Presentation1.pptx
    roundtrip ships at 4.5× output size with 4 parts removed
    (`docProps/app.xml`, `docProps/core.xml`, `ppt/tableStyles.xml`,
    `ppt/viewProps.xml`), 5 parts added (notes master, notes slide,
    extra theme variant), 21 parts changed (every slide layout,
    master, theme, presentation, top-level rels). The classifier
    tags this as `theme_loss + master_loss + style_loss +
    metadata_churn + structural_normalization + content_preserved_lossy`.
  - **Two classifier layers.** `classify_loss(...)` runs cheap
    list-based rules over the diff's part-name set;
    `classify_xml_loss(base_path, head_path, ...)` opens the
    packages and reads slide XML bytes for content-aware rules
    (Phase 1 only rule: `defaults_inlined`). `observe()` unions
    the two so callers don't need to know about the split.
  - **Resolution limit documented.** The oracle is file-level: it
    observes what GSuite emits in OOXML, not what Slides actually
    does at runtime. Semantic questions (does inheritance still
    bind? do animations play? do master edits propagate?) require
    a behavioral oracle — driving Slides UI via Playwright —
    deferred to a future spec (provisional: 032).
  - **Tests** — 25 always-on tests (classifier rules, observation
    shape, dispatcher routing, CLI parsing, mocked Drive client end-
    to-end, content-aware `defaults_inlined` detection on synthetic
    pptx pairs) + 1 gated real-network smoke test that skips unless
    `GSUITE_ORACLE_CREDS`, `GSUITE_ORACLE_SUBJECT`, and
    `GSUITE_ORACLE_FOLDER_ID` are all set.
- Phase 2 (DOCX, XLSX) and Phase 3 (ODF) are out of scope for this
  release; the orchestrator and `LossClass` are structured to
  parameterize cleanly.

## [0.7.4] - 2026-04-30

### Changed
- Word window primitives consolidated into
  `openxml_audit.docx.osa` (Spec 030). The 22 functions, constants,
  and aliases that historically lived in `tools/oracle/word_window.py`
  (since Spec 011 / 0.5.0) now live in the in-package layer parallel
  to `pptx.osa` (since 0.6.4) and `xlsx.osa` (since 0.6.5). The
  four-format oracle ladder finally has one consistent in-package
  osa shape:
  - **`docx.osa` now exports**: `WORD_APP_BUNDLE` /
    `WORD_APP_ID` / `WORD_PROCESS_NAME` /
    `REPAIR_DIALOG_PATTERNS` /
    `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS` (constants);
    `launch_word` / `launch_word_app` (alias) / `is_word_running`
    / `word_version` (lifecycle); `open_document` /
    `list_open_document_names` / `is_document_open` /
    `activate_document` (document operations); `save_document` /
    `close_document` / `close_active_document` (alias) /
    `close_document_saving` / `close_active_document_saving`
    (alias) (save/close); `find_repair_dialog_text` /
    `click_dialog_button` / `dismiss_repair_dialog` /
    `dismiss_any_leftover_modal` (dialog handling).
  - **`tools/oracle/word_window.py` is now a back-compat shim**
    that re-exports from `docx.osa` plus the shared
    `openxml_audit.osa` primitives (`osascript`, `osascript_jxa`,
    `applescript_quote`). The historical `OsascriptError` is
    aliased to `RuntimeError` so existing
    `except word_window.OsascriptError` clauses keep catching
    what they always caught (the new `openxml_audit.osa.osascript`
    raises plain `RuntimeError`).
- Future code should import from `openxml_audit.docx.osa`
  directly. The shim is permanent for back-compat through the
  remaining 0.7.x and 0.8.x releases; eventual removal is
  pre-1.0 work.

### Added
- 11 new tests in `tests/test_docx_osa.py` lock:
  - the symbol set the in-package layer must expose
  - the back-compat shim re-exports those same symbols
  - dialog-pattern shape (`unable to read this document`
    must be in the pattern list; `Open and Repair` must be in
    the accept-button list — both critical fixes from the 0.6.6
    baseline run that closed Word's hard-error dialog gap)
  - `OsascriptError` alias catches `RuntimeError`
  - `_applescript_quote` private alias is reachable for legacy
    callers
- `specs/030-word-oracle-osa-consolidation.md` documents the
  consolidation + the deferred consumer-migration follow-ups
  (`word_roundtrip.py` / `word_repair_corpus.py` /
  `word_repair_oracle.py` should eventually import from
  `docx.osa` directly; current release keeps the shim
  transparent).

### Verified
- Full test suite passes: 617 tests (up from 606 in 0.7.3, +11
  new). The matrix-driven Word oracle (Spec 010/011) continues
  to work through the shim.

## [0.7.3] - 2026-04-30

### Added
- First Excel canonical-form validator (Spec 029, Phase 1).
  Detects patterns Excel will silently rewrite on save — even
  when the file passes our schema/semantic validation cleanly.
  Triggered by the v0.7.2 baseline finding that every TokenMoulds-
  emitted `.xlsx` came back with 10 changed parts and no repair
  dialog.
  - **`src/openxml_audit/excel/canonical_form.py`** —
    `ExcelCanonicalFormValidator` class. Wired into the main
    validator's `_validate_spreadsheet_structure`.
  - **First check: `Excel_InlineStrCells`.** Flags worksheets
    with `<c t="inlineStr">` cells when the package has no
    populated `xl/sharedStrings.xml`. This is the dominant cause
    of the v0.7.2 silent-canonicalization finding: Excel migrates
    every inline string to a freshly-created shared-strings table
    on save. Severity WARNING, `source_class=EXCEL_APP_COMPAT`,
    one finding per affected worksheet with cell count in the
    description.
  - 7 new tests in `tests/test_excel_canonical_form.py` cover
    the positive cases (no SST + inlineStr cells, empty SST +
    inlineStr cells, both v0.7.2 corpus fixtures), the negative
    cases (no inlineStr, populated SST with `<c t="s">` refs,
    inlineStr with populated SST — locks the scope decision),
    and the v0.7.2 corpus smoke (`acme-us.xlsx` and
    `globex-gb.xlsx` flagged with count 9 each, matching what
    the manual probe found).
- `specs/029-excel-canonical-form-validators.md` documenting
  Phase 1 + the deferred Phase 2 (chart externalLinks
  materialization detection) and Phase 3 (non-canonical
  attribute order).

### Verified
- Self-parity against the v0.7.1 baseline: **zero drift.** The
  new check only fires on inlineStr-without-SST patterns; the
  SDK seed corpus doesn't have that shape, so the v0.7.1
  baseline passes against itself with the new check active. No
  baseline refresh needed.

## [0.7.2] - 2026-04-29

### Added
- TokenMoulds-API corpus helper (Spec 028 — Phase 2 of Spec 026's
  roadmap to 0.8.0). Resolves the corpus-curation gap flagged in
  0.6.9: instead of globbing TokenMoulds' filesystem (which mixed
  release-quality emitter output with troubleshooting-session
  leftovers), `tools/oracle/build_corpus.py` *generates* a clean
  corpus by driving TokenMoulds' Python API directly.
  - **`tools/oracle/build_corpus.py`** — calls each emitter's
    `build_package()` and post-processes the bytes via
    `template_to_document(bytes, target=<format>)` to flip the
    content-type / mimetype from template (`.dotx`, `.xltx`, etc.)
    to document (`.docx`, `.xlsx`, etc.). Lossless: only the
    format-identification bytes shift, the underlying XML is
    unchanged.
  - **Corpus** under `data/corpus/tokenmoulds_v0.7.2/` — 12 files
    (2 brand variants × 6 formats: Word, Excel, PowerPoint, ODT,
    ODS, ODP). Byte-reproducible from the same brand inputs
    run-to-run.
  - **Re-baselines** under
    `tools/oracle/baselines/<format>/v0.7.2-clean.json` for all
    four oracles. Replaces the noisy 0.6.9 baselines as the
    headline reference.
- `specs/028-tokenmoulds-corpus-helper.md` documenting design +
  the corpus-curation problem the previous run flagged.

### Documented
- `tools/oracle/baselines/README.md` extended with a "v0.7.2 clean
  run" section. Headline findings:
  - **Word and PowerPoint accept TokenMoulds output as canonical**
    — 2/2 each, 3 seconds each, zero repair dialogs. Confirms
    the 0.6.9 high-friction cases were the *corpus*, not the
    emitter.
  - **Excel rewrites every TokenMoulds-emitted workbook** — 2/2
    files came back with 10 changed parts each. No repair dialog;
    Excel chose silent canonicalization. Mission-relevant
    finding: the validator could detect what makes TokenMoulds'
    Excel output non-canonical so users avoid the issue.
  - **External-source dialog observed** — Excel showed a "This
    workbook contains links to one or more external sources..."
    modal on one of two `.xlsx` files. Different from a repair
    dialog (data-source-trust prompt, not corruption recovery);
    needs a separate detection path. Cataloged as an operational
    follow-up.
- Side-by-side comparison vs the 0.6.9 mixed-corpus baseline
  shows the corpus-curation gap is fixed: Word and PPTX
  success rates jumped from 50%/60% → 100%/100%.

### Note on self-parity baseline
Spec 026's 0.7.2 step originally proposed re-running the
self-parity baseline against the new corpus too. On reflection,
that's wrong — the self-parity baseline's purpose is to detect
**validator-output regressions**, which requires a *stable*
corpus. The SDK seed corpus (pinned to v3.4.1) is that. The
TokenMoulds corpus is for **oracle / app-survival** work and
changes as TokenMoulds evolves. The 0.7.1 self-parity baseline
against the SDK seed corpus stays as the self-parity reference.

## [0.7.1] - 2026-04-29

### Added
- Self-parity prototype as advisory CI (Spec 027 — Phase 1 of
  Spec 026's roadmap to 0.8.0). The blocking sovereign gate's
  substance lands here; 0.8.0 promotes via policy.
  - **`scripts/parity/run_self_parity_snapshot.py`** — walks the
    corpus, runs the validator at each `FileFormat`, emits a
    `family_key`-keyed inventory tagged with `SourceClass`. The
    initial baseline captured 18,516 findings across 1,368 unique
    family_keys (228 of them `word_app_compat` from Spec 010).
  - **`scripts/parity/compare_self_parity.py`** — diffs current
    vs baseline. Three threshold knobs (`--max-new-families`,
    `--max-missing-families`, `--max-count-drift-total`),
    strict-no-drift defaults. Exits nonzero on threshold
    violation; renders a markdown summary for workflow surfaces.
  - **`data/corpus/self_parity_baseline/v0.7.1/snapshot.json`**
    — initial baseline at this release's `main` HEAD.
  - **`.github/workflows/self-parity-gate.yml`** — runs on every
    push/PR, advisory only (`continue-on-error: true` on the
    comparison step). Coexists with the SDK `parity-gate.yml`.
    Two parallel informational signals; nothing blocking.
  - 10 new tests in `tests/test_self_parity.py` covering compare
    logic, threshold flags, markdown rendering, and a baseline-
    vs-self schema-shape smoke.
- `specs/026-self-parity-sovereign-gate-roadmap.md` — umbrella
  spec for the four-step path (0.7.1 → 0.7.2 → 0.7.3 → 0.8.0).
- `specs/027-self-parity-prototype.md` — design + Phase 2
  deferrals.

## [0.7.0] - 2026-04-29

### Cutover release

This is the consolidation of the 0.6.x staircase. The substantive work
shipped in 0.6.1 → 0.6.9; 0.7.0 is the moment to acknowledge it as a
coherent feature set with one notable user-visible polish — the
`openxml-audit-oracle` console script — and a SemVer minor bump to
mark the new shape of the validator.

**Where 0.7.0 leaves the project:**

- **All four formats have a roundtrip oracle** that drives the target
  app, observes save behavior, and emits a structured observation
  report. Word (Spec 011, since 0.5.0), ODF (Spec 019, 0.6.3),
  PowerPoint (Spec 020, 0.6.4), Excel (Spec 021, 0.6.5).
- **Per-part canonical-c14n diffs** in the observation reports
  (Spec 024, 0.6.8) — replaces the hash-only diff that earlier
  releases shipped on three of the four oracles. PPTX's existing
  per-part text diff was extracted into the new shared
  `openxml_audit.package_diff` module so XLSX, ODF, and Word now
  emit the same shape.
- **Auto-dismiss for repair dialogs** in PPTX and XLSX
  (Spec 023, 0.6.7) — closes the symmetry gap with Word's oracle.
  ~12× speedup in roundtrip wall-clock time on files that trigger
  Office's repair flow.
- **Source-class tagging on `ValidationError`** (Spec 018, 0.6.2)
  — `SourceClass.SDK_PROXY` / `WORD_APP_COMPAT` / `EXCEL_APP_COMPAT`
  / `POWERPOINT_APP_COMPAT` / `ODF_NATIVE` / `PACKAGE_INTEGRITY`,
  exported at package top level. Consumers (the future Spec 013
  sovereign gate, downstream pytest plugins) can filter findings
  by source.
- **Path-indexing fix** so error paths match the SDK's XPath
  conventions (Spec 014, 0.6.1). On
  `Document.docx` Office2007, the parity gate's "+1 mystery" from
  Spec 012 reduced from 50+ phantom family deltas to a single real
  structural difference.
- **macOS preflight + permissions documentation** (Spec 023, 0.6.5).
  `python -m tools.oracle.preflight` checks all four engines;
  `docs/oracle_permissions.md` is the setup checklist.
- **Two committed baseline runs** under `tools/oracle/baselines/`
  (Specs 022 & 025, 0.6.6 & 0.6.9). The validator has *evidence*,
  not just tools.

### Added
- New console script: `openxml-audit-oracle <engine> FILES...`
  Replaces the longer `python -m openxml_audit.oracle <engine>`
  introduced in 0.6.8 — same dispatcher, friendlier name, on-PATH
  after `pip install`. Engines: `word` / `excel` (alias `xlsx`) /
  `pptx` (alias `powerpoint`) / `odf` / `preflight`.

### Deferred to future releases (explicitly tracked)

- **Self-parity sovereign gate** (Spec 013) — the eventual
  blocking gate consumes oracle output across all four formats.
  0.7.x or 0.8.x candidate.
- **Word oracle migration to in-package `osa` layer** —
  `tools/oracle/word_window.py` and `word_roundtrip.py` predate the
  `src/openxml_audit/docx/osa.py` consolidation pattern and still
  duplicate primitives. Worth folding before 0.8.0 but the
  matrix-driven Word oracle (Spec 010 / 011) depends on the
  current layout, so this is its own focused release.
- **Clean re-baseline against fresh TokenMoulds-API output**
  (Spec 025 Phase 2) — the corpus-curation problem flagged in the
  0.6.9 baseline. Needs a small `tools/oracle/build_corpus.py`
  helper that drives TokenMoulds emitters programmatically to
  produce concrete `.docx`/`.xlsx`/`.pptx`/`.odt` (not templates).
- **Repair categorization on top of per-part diffs** (Phase 2 of
  Specs 019/020/021/022) — distinguish cosmetic XML reflow from
  substantive content changes.
- **Pattern-list expansions** for dialog wordings observed during
  the 0.6.9 walk (Word "Open and Repair" button-label edge case,
  PowerPoint "[Repaired]" secondary modal). Cataloged in
  `tools/oracle/baselines/README.md`.

## [0.6.9] - 2026-04-29

### Added
- Second baseline run across all four formats (Spec 025), 4× larger
  corpus than the first baseline (0.6.6). Written under
  `tools/oracle/baselines/{word,odf,pptx,xlsx}/2026-04-29-larger.json`.
  Made tractable by 0.6.7's auto-dismiss; annotated with the per-part
  diffs (`added_parts` / `removed_parts` / `diff_dir`) shipped in 0.6.8.
  First release-grade use of the meta-CLI from 0.6.8
  (`python -m openxml_audit.oracle <engine>`).

### Documented
- `tools/oracle/baselines/README.md` extended with a "2026-04-29
  second run" section covering aggregate results, the corpus-quality
  caveat (the assembled corpus mixed release-quality TokenMoulds
  emitter output with troubleshooting-session leftovers — the
  elevated repair-dialog rate likely reflects the latter), and a
  catalog of operational follow-up patterns observed during the
  walk:
  - Word "Open and Repair" button-label match was inconsistent —
    possibly a System Events label-string variant.
  - PowerPoint "[Repaired] and removed it. View / Delete"
    secondary info modal needs a button label outside the current
    accept list (or title-bar "[Repaired]" detection + Escape).
  - Excel post-`pkill` recovery sheet (already documented from
    the first run; same benign behavior).
- `specs/025-larger-corpus-baselines.md` documenting design,
  Phase 2 deferrals (clean re-baseline against fresh TokenMoulds-
  API-driven output), and risks.

## [0.6.8] - 2026-04-29

### Added
- **Shared per-part diff module** `openxml_audit.package_diff`
  (Spec 024). Format-agnostic c14n + unified-diff machinery
  extracted from `pptx.lab` so the XLSX, ODF, and Word corpus
  oracles can replace their hash-only diff with per-part text
  diffs. Public API: `canonicalize_xml`, `load_package_parts`,
  `compare_package_parts`, `compare_packages`, `pretty_part_text`,
  `sanitize_part_name`, `write_part_diff`.
- **Meta-CLI** `python -m openxml_audit.oracle <engine> ...`
  dispatching to all four corpus oracles plus the preflight
  check. Subcommands: `word`, `excel` (alias `xlsx`), `pptx`
  (alias `powerpoint`), `odf`, `preflight`. One verb across
  formats — useful for shell history and for downstream
  callers that don't need to remember which file each oracle
  lives in under `tools/oracle/`.
- 11 new tests in `tests/test_package_diff.py` covering
  canonicalization (whitespace-only diffs collapse), custom
  `parts_filter`, parse-error fallbacks, end-to-end report
  shape, and unified-diff output.

### Changed
- XLSX, ODF, and Word corpus oracles now emit **per-part text
  diffs** instead of just hash deltas.
  - `RoundtripObservation` for all three formats gains
    `added_parts: list[str]`, `removed_parts: list[str]`, and
    `diff_dir: str | None` (None unless `--keep-artifacts`).
  - Each oracle CLI gains a `--keep-artifacts` flag. Without it,
    diff dirs are cleaned up after the run; with it, callers can
    inspect `<work_dir>/compare/diffs/<sanitized-part>.diff` to
    see what the target app actually rewrote on save.
  - Phase 1 of the per-part repair categorization Phase 2 of
    Specs 019/020/021/022 needs.
- `pptx.lab` now imports its diff primitives from
  `openxml_audit.package_diff` rather than defining them locally.
  Public API (`compare_pptx_packages`, `write_pptx_snapshot`)
  unchanged. PPTX-specific timing-tree change collector stays
  put.

### Fixed
- Latent bug in `pretty_part_text` (the function moved into the
  new shared module): with `recover=True` set on the parser,
  malformed XML returns `None` from `etree.fromstring` rather
  than raising — the code path then crashed in `etree.tostring`.
  Now falls back to raw decoded text. This bug existed in
  `pptx.lab._pretty_part_text` since at least 0.5.0 but was
  unreachable in PPTX practice (PPTX inputs are well-formed).
  The new test suite surfaced it.

## [0.6.7] - 2026-04-29

### Added
- Auto-dismiss for PowerPoint and Excel repair dialogs (Spec 023,
  Phase 1). Closes the symmetry gap with Word — Word's oracle has
  had `click_dialog_button` since Spec 011; PPTX and XLSX now have
  it too. The 0.6.6 baseline run on `winsemius.tokens.xlsx` took
  61.4 seconds with manual click; the same file roundtrips in
  ~5 seconds with auto-dismiss. Concretely:
  - `pptx.osa.click_dialog_button(label)` and
    `xlsx.osa.click_dialog_button(label)` — System Events button
    click by exact label, mirroring `tools/oracle/word_window.click_dialog_button`.
  - `pptx.osa.dismiss_repair_dialog()` and
    `xlsx.osa.dismiss_repair_dialog()` — high-level helper:
    detect via `find_repair_dialog_text`, try each label in
    `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`, return
    `(was_seen, dialog_text, clicked_label)`. `clicked_label`
    is None when the dialog appeared but no button matched —
    the orchestrators bail with `open_failed` in that case.
  - `pptx.osa.dismiss_any_leftover_modal()` and
    `xlsx.osa.dismiss_any_leftover_modal()` — sends Escape
    (key code 53) to the frontmost app process. Used as a
    finalize-after-close step to clear secondary info modals
    that don't have an obvious accept button (e.g. Excel's "we
    made repairs — View / Delete" sheet).
  - `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS` — per-app accept-button
    preference lists (PPTX leads with `Repair`, XLSX leads with
    `Yes`).
  - `RoundtripObservation.repair_dialog_button_clicked` — new
    field on both PPTX and XLSX observations recording which
    label cleared the dialog. None when no dialog was seen.
  - Orchestrator changes in `tools/oracle/{pptx,xlsx}_repair_oracle.py`:
    auto-dismiss runs interleaved with the open-poll loop, after
    open registers, and after close-with-save (Escape cleanup).
- 4 new tests in `tests/test_pptx_roundtrip_oracle.py` and
  `tests/test_xlsx_roundtrip_oracle.py` covering the new exports,
  preference-list shape, and dismiss-helper return contract on
  no-dialog state.

### Fixed
- `pptx.osa.find_repair_dialog_text` and
  `xlsx.osa.find_repair_dialog_text` now catch
  `subprocess.TimeoutExpired` (in addition to `RuntimeError`),
  matching the same fix applied to `list_open_*_names` in 0.6.5.
  Without this, calling the function with the target app cold
  raised an unhandled timeout.

## [0.6.6] - 2026-04-29

### Added
- First oracle baselines committed across all four formats (Spec 022,
  Phase 1). Real observations from running the oracles on
  TokenMoulds-emitted corpora — moves the project from "we built
  four oracle tools" to "we ran them and have evidence."
  - `tools/oracle/baselines/odf/2026-04-29.json` — 5 files, all
    `repaired` (soffice canonicalizes formatting on every save).
  - `tools/oracle/baselines/word/2026-04-29.json` — `mcp-word.docx`
    came back `preserved` after Word's repair dialog fired (caught
    and dismissed by the new patterns added below).
  - `tools/oracle/baselines/pptx/2026-04-29.json` — 2 files,
    `winsemius.pptx` and `mcp-swot.pptx` both `preserved` with no
    repair dialogs.
  - `tools/oracle/baselines/xlsx/2026-04-29.json` —
    `winsemius.tokens.xlsx` triggered Excel's "found a problem" repair
    dialog; post-recovery canonical XML matched the input byte-for-byte.
  - `tools/oracle/baselines/README.md` — methodology, results table,
    per-format findings, operational notes (Excel `.xltx` template
    limitation, post-`pkill -9` recovery prompts, etc.).
  - `specs/022-first-oracle-baselines.md` — design + Phase 2 deferrals.
- New `tools/oracle/word_repair_corpus.py` — sibling of the ODF/PPTX/
  XLSX corpus walkers. The existing `word_repair_oracle.py` is a
  scenario-matrix runner (Spec 010/011 phase work); this new tool
  walks arbitrary `.docx` corpora using `word_roundtrip.roundtrip()`
  and emits the same `RoundtripObservation` shape as the other three
  oracles. CLI:
  `python tools/oracle/word_repair_corpus.py FILES... [--output X]`.

### Fixed
- `pptx.osa.list_open_presentation_names` and
  `xlsx.osa.list_open_workbook_names` switched their AppleScript
  enumeration from `repeat with X in COLLECTION / name of X` to
  `name of every X`. The broken idiom hangs Office for Mac M365's
  AppleScript engine indefinitely (Word's oracle already knew this
  and used the correct form since Spec 011; PPTX and XLSX inherited
  the wrong pattern in 0.6.4 / 0.6.5). Symptom before the fix:
  oracle would time out at the polling-window deadline because the
  enumeration helper kept returning `[]` even with files actually
  open in the target app.
- `tools/oracle/word_window.REPAIR_DIALOG_PATTERNS` extended with
  the "file is corrupt → Open and Repair?" dialog wording (`unable
  to read this document`, `may be corrupt`, `open and repair`,
  `text recovery converter`). Without these patterns, the Word
  oracle hung when Word presented its hard-error dialog instead of
  the soft repair dialog.
- `tools/oracle/word_roundtrip._try_dismiss_repair_dialog`'s
  alt-button label list extended with `("Open and Repair",
  "Cancel")` so the oracle clicks through the new dialog
  variant correctly.

## [0.6.5] - 2026-04-29

### Added
- Excel roundtrip oracle (Spec 021, Phase 1). Fourth and final entry
  in the oracle ladder — the validator now has a roundtrip oracle
  for every supported format (Word + ODF + PowerPoint + Excel).
  - **Extends `src/openxml_audit/xlsx/osa.py`** with the primitives
    that already existed in `docx.osa` and `pptx.osa` but were
    missing from `xlsx.osa`: `list_open_workbook_names`,
    `is_workbook_open`, `close_workbook_saving`,
    `find_repair_dialog_text`, and the public
    `REPAIR_DIALOG_PATTERNS` match list.
  - **`tools/oracle/xlsx_repair_oracle.py`** is the orchestrator.
    Stages input under `~/Documents/.xlsx_oracle_runs/<id>/` (Excel
    App Sandbox default), uses the existing `xlsx.osa` layer for
    window control, fingerprints canonical OOXML parts before/after
    (workbook.xml, sheets/, styles, sharedStrings, theme/, charts/,
    pivots, etc.), and emits the same `RoundtripObservation` shape
    as the Word/PowerPoint/ODF oracles. CLI:
    `python tools/oracle/xlsx_repair_oracle.py FILES... [--output X]`.
  - 11 tests in `tests/test_xlsx_roundtrip_oracle.py` (9 always-on
    + 2 Excel-required, skip cleanly when Excel isn't installed).

### Changed
- **Generalized `tools/oracle/preflight.py`** to cover Word + Excel
  + PowerPoint + LibreOffice. The original Word-only `check()` is
  kept as a back-compat alias for `check_word()`. Run all engines
  with `python -m tools.oracle.preflight` or one with
  `--engine <name>`.
- `xlsx.osa.list_open_workbook_names` and
  `pptx.osa.list_open_presentation_names` now catch
  `subprocess.TimeoutExpired` (in addition to `RuntimeError`) so a
  cold-launch of the underlying Office app no longer hangs the test
  suite past the 10s budget. Symptom before the fix: the smoke test
  would timeout on first run if Excel/PowerPoint weren't already
  open. Returns `[]` on timeout — same as on permission denied.

### Documentation
- New `docs/oracle_permissions.md` — macOS setup checklist for
  the roundtrip oracles. Covers Automation grants (control Word,
  Excel, PowerPoint via AppleScript), Accessibility grants (System
  Events keystrokes for `Cmd-S` save paths), App Sandbox staging
  directories, the symptom checklist for denied / revoked
  permissions, and `python -m tools.oracle.preflight` as the
  canonical readiness check.

## [0.6.4] - 2026-04-29

### Added
- PowerPoint roundtrip oracle (Spec 020, Phase 1). The orchestrator
  is the third in the family — sibling of the Word oracle (Spec 011)
  and the ODF oracle (Spec 019, 0.6.3). Deliberately built on top of
  the existing PPTX layers rather than duplicating them:
  - **Extends `src/openxml_audit/pptx/osa.py`** with the primitives
    that were already present in `docx.osa` and `xlsx.osa` but missing
    from PowerPoint: `list_open_presentation_names`,
    `is_presentation_open`, `save_presentation`, `close_presentation`,
    `close_presentation_saving`, `find_repair_dialog_text`, and the
    public `REPAIR_DIALOG_PATTERNS` match list.
  - **`tools/oracle/pptx_repair_oracle.py`** wires the existing
    `pptx.osa` window control + `pptx.lab.compare_pptx_packages`
    per-part diff into the standard `RoundtripObservation` shape used
    by Word and ODF. Stages the input under
    `~/Documents/.pptx_oracle_runs/<id>/` (PowerPoint App Sandbox
    grants access there by default; `PPTX_ORACLE_STAGE` env var
    overrides). CLI:
    `python tools/oracle/pptx_repair_oracle.py FILES... [--output X]`.
  - 7 new tests in `tests/test_pptx_roundtrip_oracle.py` (5
    always-on + 2 PowerPoint-required, skipped cleanly when
    PowerPoint or osascript is unavailable).
  - `specs/020-pptx-roundtrip-oracle.md` documenting the design,
    the existing layers it composes (PPTX timing oracle, PPTX
    package differ, oracle starter decks), and deferred Phase 2
    work (TokenMoulds corpus, repair categorization, pairwise
    scenario mutations).

  Phase 1 detects PowerPoint's repair dialog but does not yet
  auto-dismiss it (the Word oracle has the button-click ceremony;
  PowerPoint analog deferred to Phase 2).

## [0.6.3] - 2026-04-29

### Added
- ODF roundtrip oracle infrastructure (Spec 019, Phase 1). Sibling of
  the existing Word roundtrip oracle, but uses LibreOffice's headless
  `soffice --convert-to` rather than osascript-driven Microsoft Word.
  - `tools/oracle/odf_window.py` — crash-resistant soffice harness.
    Each call gets a fresh `UserInstallation` profile dir (no
    single-instance lock contention), runs under a hard wall-clock
    timeout, and reaps descendant helper processes (oosplash etc.)
    via pgrep + SIGKILL on the process group when the timeout fires.
    Returns a structured `SofficeRunResult` rather than raising on
    soft errors, so corpus walks record "this file crashed soffice"
    as data instead of stalling the run.
  - `tools/oracle/odf_repair_oracle.py` — observation layer. For each
    input ODF, fingerprints the canonical parts (`content.xml`,
    `styles.xml`, `meta.xml`, `settings.xml`) before and after the
    soffice roundtrip, classifies the outcome as `preserved` /
    `repaired` / `crash` / `timeout` / `open_failed` / `missing_output`,
    and emits a JSON observation report. CLI:
    `python tools/oracle/odf_repair_oracle.py FILES... [--output X]`.
  - 8 new tests in `tests/test_odf_roundtrip_oracle.py` (6 always-on
    pure-logic tests + 2 soffice-required integration tests that skip
    cleanly when no LibreOffice is installed).
  - `specs/019-odf-roundtrip-oracle.md` documenting the design and
    deferred phases (corpus generation via TokenMoulds, structural
    XML diff, pairwise scenario mutations).

  Phase 1 only ships the harness and observation skeleton. Phase 2
  (0.6.4) adds the TokenMoulds-driven corpus + structural diff that
  distinguishes cosmetic repairs from substantive ones.

## [0.6.2] - 2026-04-29

### Added
- New `SourceClass` enum on `ValidationError` (also exported at package
  top level). Six values: `sdk_proxy`, `word_app_compat`,
  `excel_app_compat`, `powerpoint_app_compat`, `odf_native`,
  `package_integrity`. Lets parity tooling separate findings that
  mirror the .NET Open XML SDK from app-survival findings unique to
  this validator. Foundational for Spec 013's self-parity gate
  (where the SDK signal needs to be cleanly filterable). Default is
  `SDK_PROXY`; explicit emission sites updated:
  - `WordCompatValidator` (Spec 010 child-ordering): `WORD_APP_COMPAT`.
  - `StylesWithEffectsValidator` consistency checks (the python-docx
    repair-dialog failure mode): `WORD_APP_COMPAT`. Schema/structural
    checks of the same part stay `SDK_PROXY` since the SDK validates
    them too.
  - PPTX `presProps`/`viewProps`/`tableStyles` missing-relationship
    finding: `POWERPOINT_APP_COMPAT`.
  - `OdfValidator._create_result` tags every finding with default
    `SDK_PROXY` as `ODF_NATIVE` at the boundary, so future ODF
    emit sites are correctly classified without per-call work.
- 6 new tests in `tests/test_source_class_tagging.py` lock the
  contract: enum values, default tagging, propagation through
  `ValidationContext`, public re-export.

## [0.6.1] - 2026-04-29

### Fixed
- Validation error paths now include 1-indexed sibling-position predicates
  (e.g. `/document[1]/body[1]/sdt[5]/...`), matching the .NET Open XML SDK's
  `XPathFinder` output. Previously every segment was emitted without a
  position predicate and the parity normalizer blindly appended `[1]`,
  collapsing distinct elements with the same tag into a single
  `family_key`. Verified against the live SDK runtime via
  `tools/parity/dotnet_validator_runner` on `TestFiles/Document.docx`:
  Office2007 family deltas dropped from 50+ to zero. The remaining
  Spec-012 "+1 mystery" is now traceable to a single structural
  divergence (our validator flattens through `<w:sdt>` wrappers; SDK
  preserves them — separate follow-up).

## [0.6.0] - 2026-04-29

### Breaking
- `stylesWithEffects` consistency check is now ERROR (was WARNING). Files
  that previously validated with this warning will now fail validation.
  This matches the severity of the reverse-direction check and reflects
  the Word-application impact (the "unreadable content" repair dialog).

### Added
- `stylesWithEffects` consistency validator now flags styles present in
  `styles.xml` but missing from `stylesWithEffects.xml` (the python-docx
  failure mode that produces Word's "unreadable content" repair dialog),
  and flags differing or one-sided `w:docDefaults` between the two parts.
- PPTX presentation validator now flags missing `presProps`, `viewProps`,
  and `tableStyles` relationships as ERROR. PowerPoint triggers its
  "unreadable content" repair dialog when these parts are absent, even
  when the package is internally self-consistent (relationship and
  content-type entries removed too). ECMA-376 makes them optional, but
  this is empirically required by PowerPoint.
- Word compatibility ordering check, Phase 1 + Phase 2. Flags child
  element reorderings inside WordprocessingML property complex types
  that trigger Word's "unreadable content" repair dialog despite the
  .NET Open XML SDK accepting the same files. Severity WARNING — this
  is an empirical Word-app-compat finding, not a strict OOXML violation.
  - Phase 1 covers `CT_TrPr` (e.g. `cantSplit` after `tblHeader`,
    issue #3's repro).
  - Phase 2 covers `CT_TblPr`, `CT_TcPr`, `CT_SectPr` — corpus-validated
    against the TokenMoulds template-generator output (33,554 property
    subtrees, 100% pass against the SDK proxy).
  - `CT_PPr` and `CT_RPr` are deferred to Phase 3: empirical mining of
    the same corpus shows the SDK proxy is too strict for `rPr` (~1.4%
    deviation rate) and so cannot ship without a corpus-derived
    canonical ordering.
- New mining tool `scripts/mine_word_property_orderings.py` and docs at
  `docs/word_compat/` — given a DOCX corpus, produces an empirical
  ordering report that validates or refutes the SDK proxy per type.
- Word roundtrip oracle infrastructure (`tools/oracle/`, spec 011
  Phase 1). Developer-machine tool that opens a DOCX in Microsoft Word
  for Mac, saves it back through Word, and exposes the input/output diff
  as ground truth for "would Word repair this?" questions. Stages
  files under `~/Documents/.word_oracle_runs/` (Word's App Sandbox
  blocks tmpfs paths). The close-with-save persist path works; Word's
  bespoke `save as` and the inherited Cocoa `save` both return -1708
  despite their dictionary declarations.
- Spec 010 oracle driver (`tools/oracle/word_repair_oracle.py`, spec 011
  Phase 2). Generates a DOCX scenario matrix using python-docx (real
  Word-template foundation) with property-element children mutated to
  test specific orderings, roundtrips each through Word, and emits a
  JSON baseline at `tools/oracle/baselines/`.
- Oracle scenario driver generalized over a `MATRICES` registry covering
  CT_TrPr (12 children), CT_TblPr (14 children), CT_TcPr (13 children),
  and CT_SectPr (18 children). Each matrix carries its own canonical
  child order, minimum-valid attribute set, and host materializer (the
  python-docx scaffold the property element is embedded in). New CLI:
  `python -m tools.oracle.word_repair_oracle {trpr|tblpr|tcpr|sectpr}`.
- First committed oracle baselines, all run on Word for Mac M365
  16.89.1 (`tools/oracle/baselines/`):
  - `word_trpr_pairwise.json`: 68 scenarios, 68 preserved.
  - `word_tblpr_pairwise.json`: per the matrix run.
  - `word_tcpr_pairwise.json`: per the matrix run.
  - `word_sectpr_pairwise.json`: per the matrix run.

  All four matrices include the baseline (canonical order) as a control,
  every pairwise child swap, and the canonical fully reversed. Issue
  #3's specific `CT_TrPr` `tblHeader`-before-`cantSplit` pattern is one
  of the 68 preserved cases. The `CT_TblPr`/`CT_TcPr`/`CT_SectPr`
  constraints shipped earlier in this release were corpus-validated;
  the oracle baselines either confirm or refute that signal directly.
  Spec 010 Phase 1's `CT_TrPr` constraint is now empirically refuted
  on this Word build. WARNING severity (not ERROR) was clearly the
  right call across all of these constraints; some may be removed in
  a future release once additional Word builds are surveyed.

### Changed
- **SDK parity gate is now advisory** (Spec 012). The
  `.github/workflows/parity-gate.yml` "Compare against baseline" step
  carries `continue-on-error: true`. SDK drift is surfaced in the
  workflow summary for trend visibility but no longer fails the build.
  The perf-budget guard remains blocking. Branch protection on `main`
  must be updated separately (manual GitHub UI step) to remove the
  parity gate from required status checks.
- `data/corpus/sdk_seed/manifest.json` is now mixed-semantics: most
  entries are SDK-extracted expectations, but four entries on
  `TestFiles/Document.docx` are manually adjusted to match this
  validator's output (annotated with `adjusted_for_app_compat`). The
  baseline reports 100% match rate against the adjusted manifest.
  Spec 013 will formalize the self-parity vs SDK-parity split.
- `.github/workflows/calibrate-parity.yml` now runs the snapshot step
  against the committed manifest rather than the freshly-extracted
  runtime manifest. Eliminates a check-universe divergence between
  calibration and the parity gate (autoplan codex finding #6).

- .NET SDK runtime parity tool (`tools/parity/dotnet_validator_runner/`
  + `scripts/parity/diff_sdk_runtime.py`). Spec 013 OQ8(B) prototype
  that invokes `OpenXmlValidator` directly against the corpus and
  diffs against our Python validator's output. The first run already
  surfaced concrete signal: the historical "+1" on `Document.docx`
  (the surprise the autoplan flagged) is a path-element-indexing
  discrepancy between Python and SDK on `<w:sdt>` paths, not four
  unrelated mysterious findings. Not yet wired into CI; tool only.

### Deferred
- Sovereign blocking gates (self-parity + Word/PowerPoint/Excel
  roundtrip oracle) deferred to Spec 013
  (`specs/013-validator-output-sovereign-gates.md`). Stub committed.
  Until Spec 013 lands, only the perf budget is blocking.

## [0.5.0] - 2026-04-20

### Added
- committed PPTX oracle deck scaffolds under `data/pptx_oracle/scaffolds/`,
  shipped inside the wheel so oracle starter decks can be materialized from
  packaged assets instead of scratch-only local files
- scaffold materialization helpers for committed PPTX oracle packages, with
  coverage asserting the wheel contains the expected PowerPoint package parts

### Changed
- `build_oracle_starter_deck()` and `build_timing_oracle_deck()` now rebuild
  decks from scaffold package trees instead of requiring `python-pptx` at
  runtime
- PPTX oracle builder and lab modules now target the shipped scaffold data as
  their runtime source of truth, while keeping the maintenance path available
  for scaffold regeneration
- repository linting was normalized so `ruff check .` passes again across the
  shipped package, scripts, and tests

## [0.4.3] - 2026-03-12

### Added
- pytest plugin with auto-registered fixtures (`assert_valid_pptx`,
  `assert_valid_docx`, `assert_valid_xlsx`, `assert_valid_odf`) — zero
  config, just `pip install openxml-audit` and use in tests
- GitHub Action for validating Office files in PRs (`changed-only` mode)
- Pre-commit hook for OOXML and ODF files
- Example scripts for python-pptx, openpyxl, ODF, and CI batch validation

### Changed
- 2x batch validation speedup via hot-path optimizations:
  single-candidate constraint cache, ignorable namespace passdown,
  singleton particle validators, and XML parse caching
- Warm validation of Document.docx (798K): 164ms → 101ms

## [0.4.2] - 2026-03-11

### Fixed
- 100% parity with Open XML SDK v3.4.1 (up from 98.7%) — 77/77 checks matched
  - 1 false-positive `Sem_UniqueAttributeValue` error for Document.docx at
    Office 2010 eliminated via version-aware MCE resolution in semantic validator
- Removed parity waiver (no longer needed)

## [0.4.1] - 2026-03-11

### Added
- Version-aware attribute gating: attributes introduced in later Office versions
  are flagged as undeclared when validating against earlier formats
- Unique attribute value validation (`Sem_UniqueAttributeValue`) for VML elements
- VML attribute namespace fix in schematron bridge (unnamespaced attributes
  were incorrectly resolved to the element's namespace)

### Fixed
- Parity improvement from 93.51% to 98.7% against Open XML SDK v3.4.1
  - 414 undeclared attribute errors now detected for Document.docx at Office2007
  - 33 undeclared attribute errors now detected for complex2010.docx at Office2007
  - 1 undeclared `shapeId` attribute error now detected for Spreadsheet.xlsx at Office2007
  - 1 `Sem_UniqueAttributeValue` error now detected for Document.docx at Office2007/2013

## [0.4.0] - 2026-03-11

### Added
- Version-aware element and attribute gating: elements introduced in later
  Office versions (e.g., Office 2010) are flagged when validating against
  earlier formats (e.g., Office 2007)
- Version-aware MCE resolution: `mc:Choice` branches requiring extension
  namespaces from later versions fall back to `mc:Fallback` at earlier formats
- Content model filtering by version: particle constraints from later versions
  are removed from content models during validation
- Properties validation for core (`docProps/core.xml`), extended (`docProps/app.xml`),
  and custom (`docProps/custom.xml`) package metadata
- Styles-with-effects validation for Word documents (`stylesWithEffects.xml`)

### Fixed
- False positive for `HLinks` element in extended properties validation

## [0.3.0] - 2026-03-10

### Added
- ODF semantic validation expanded from 10 to 27 rules (`ODFSEM001`--`ODFSEM027`)
- Supertheme part validation for PowerPoint (Office 2013+)
- Markup compatibility (MCE) validation phases for OOXML
- Gap-closure for ODF-OOXML validation milestones M1--M6

## [0.2.0] - 2026-03-09

### Added
- ODF (ODT/ODS/ODP) validation with staged conformance levels:
  foundation, schema-core, semantic-core, security-core
- ODF Relax NG schema-core routing with versioned schema maps
- ODF security policy validation (signatures, encryption structure)
- ODF reference calibration tooling (ODF Toolkit, OPF comparisons)
- ODF benchmarking scripts with per-phase timing breakdown
- Parity gate CI workflow enforcing SDK baseline match rate
- Parity baseline extraction, comparison, and waiver tooling
- Performance budget guard for validation timing
- Quarterly SDK update workflow for upstream tracking
- Release infrastructure (PyPI publishing, docs site)

### Changed
- Schema validation hot paths optimized with relationship caching
- Phase timing metrics added to validation pipeline

### Fixed
- Schema and relationship parity gaps aligned with SDK output
- Nested AND/OR schematron parsing

## [0.1.0] - 2026-03-08

### Added
- Initial release
- OOXML validation for PPTX, DOCX, and XLSX files
- Package structure validation (ZIP, content types, relationships)
- Schema validation (particle validators, type validators)
- Semantic validation (attribute, relationship, reference constraints)
- CLI with text, JSON, and XML output formats
- Python API with `OpenXmlValidator`, `validate_pptx`, `is_valid_pptx`
- Integration helpers: context managers, decorators, pytest fixtures
- Support for Office 2007 through Microsoft 365 format versions

[Unreleased]: https://github.com/BramAlkema/openxml-audit/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.9...v0.5.0
[0.4.3]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.2...v0.4.3
[0.4.2]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.1...v0.4.2
[0.4.1]: https://github.com/BramAlkema/openxml-audit/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/BramAlkema/openxml-audit/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/BramAlkema/openxml-audit/releases/tag/v0.1.0
