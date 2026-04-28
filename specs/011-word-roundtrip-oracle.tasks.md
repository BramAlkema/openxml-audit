# Tasks: Word Roundtrip Oracle for App-Compat Findings

**Spec:** [011-word-roundtrip-oracle.md](./011-word-roundtrip-oracle.md)

## Phase 1: Engine

### Module scaffolding

- [ ] Create `tools/oracle/__init__.py` and `tools/oracle/word_window.py`, `word_roundtrip.py` placeholders with module docstrings citing the spec
- [ ] Add `tools/oracle/README.md` with first-pass setup instructions: macOS, Microsoft 365 Word, UI scripting + Screen Recording permissions, how to grant them, how to recover from a stuck Word session
- [ ] Add `tools/oracle/preflight.py` — quick check that surfaces missing permissions and an unreachable Word app before the engine tries to run

### word_window.py (osascript helpers)

- [ ] Port the osascript / JXA helper shape from TokenMoulds' `tools/visual/pptx_window.py`, translated to Word: `osascript()`, `osascript_jxa()`, `_word_app_target()`, `launch_word_app()`
- [ ] Function: open a DOCX in Word via the AppleScript `tell application "Microsoft Word" to open file POSIX path of "..."` command
- [ ] Function: close the active Word document via `close active document saving no`
- [ ] Function: detect a "Microsoft Word" alert dialog whose text matches a configurable pattern list (e.g. `"unreadable content"`, `"recover the contents"`)
- [ ] Function: dismiss a detected dialog with either Yes or No, and return the captured text

### word_roundtrip.py (engine)

- [ ] Define `RoundtripResult` dataclass: `input_path`, `output_path`, `repair_dialog_seen`, `repair_dialog_text`, `elapsed_seconds`, `word_version`
- [ ] Implement `roundtrip(input_docx, *, output_dir=None, timeout=60.0, accept_repair=True) -> RoundtripResult`
- [ ] Stage input in a fresh temp directory before any Word operation
- [ ] Launch / reach Word app via `open -W -a "Microsoft Word"` with timeout
- [ ] Open the staged DOCX via the AppleScript `open` command
- [ ] Poll for the document's name in the `documents` collection until it appears or `timeout` expires
- [ ] Detect repair dialog after open; capture text; dismiss as Yes if `accept_repair`, No otherwise
- [ ] Trigger Word's Save: try AppleScript `save` first; fall back to Cmd+S keystroke if `save` doesn't trigger the same repair flow
- [ ] Capture Word version string into the result
- [ ] Close the active document; ensure the engine returns even if Word holds a separate stuck dialog
- [ ] Hard timeout with cleanup: force-close documents via osascript if the polled save loop never completes

### Engine validation

- [ ] Sanity test: roundtrip a known-good DOCX (built from the existing `_build_docx` helper); assert `preserved`-shaped output, no dialog seen
- [ ] Sanity test: roundtrip Shaun's exact issue #3 repro (`tblHeader` before `cantSplit`); assert `repair_dialog_seen=True` and post-Word XML differs from input
- [ ] Sanity test: roundtrip an obviously-malformed DOCX (e.g., schema violations); assert behavior is one of `accept_repair=False` failing or `accept_repair=True` succeeding with dialog text captured

### Test markers and CI

- [ ] Add `requires_word_app` pytest marker to `pyproject.toml` `[tool.pytest.ini_options]`
- [ ] Add an autouse fixture / plugin hook that skips marked tests when no Word app is reachable (re-uses the preflight check)
- [ ] Add `tests/test_word_oracle_engine.py` covering the parts that don't need Word (XML diff helper, repair-dialog text matching, scenario-id slugification)
- [ ] Add `tests/test_word_oracle_smoke.py` marked `requires_word_app`; runs one end-to-end roundtrip
- [ ] Confirm full pytest run on a machine without Word skips Word-dependent tests cleanly

## Phase 2: Spec 010 driver and CT_TrPr baseline

### Driver and scenario generation

- [ ] Create `tools/oracle/word_repair_oracle.py` with a CLI entrypoint
- [ ] Create `tools/oracle/scenarios/property_ordering.py` with scenario generators
- [ ] Generator: baseline (canonical-ordered children only) per `CONSTRAINT_TABLE` entry
- [ ] Generator: every pairwise swap among canonical children
- [ ] Generator: skip-then-restore patterns (`[c0, c2, c1]`, `[c0, c2, c3, c1]`, etc.)
- [ ] Stable scenario IDs derived from `parent_local + scenario kind + child names`
- [ ] DOCX synthesis: minimal valid DOCX with exactly one parent property element instance carrying the requested children. Reuse `_build_docx` / `_trpr` helpers from the test suite (extract to a shared `tools/oracle/scenarios/_docx_builders.py` if needed)

### CT_TrPr oracle run

- [ ] Run the oracle for `CT_TrPr`; collect verdict per scenario; write to `tools/oracle/baselines/word_trpr_pairwise.json`
- [ ] Confirm the issue #3 repro reports `repaired`
- [ ] Confirm the canonical baseline reports `preserved`
- [ ] Surface any scenarios that report `preserved` despite being out-of-order against the SDK proxy — these are over-flags in spec 010 Phase 1 and should be removed from the constraint
- [ ] Update spec 010's `CT_TrPr` constraint comment to cite the oracle baseline JSON for traceability

### Documentation

- [ ] Append a "Phase 2 oracle results" section to `docs/word_compat/README.md` with a small table summarising verdicts per scenario kind
- [ ] Update spec 010 Phase 1 description to "shipped + oracle-validated"
- [ ] Update spec 010 Phase 4 to reference the oracle as the way Phase 4 work happens (no longer waiting on Shaun's corpus)

## Phase 3: rPr and pPr deviation oracle

### Driver extension

- [ ] Generator: load corpus deviation patterns from a TokenMoulds (or other corpus) mining run via `scripts/mine_word_property_orderings.py --json`, feed each as a scenario
- [ ] Run oracle for `CT_RPr` deviation patterns; commit `tools/oracle/baselines/word_rpr_deviations.json`
- [ ] Run oracle for `CT_PPr` deviation patterns; commit `tools/oracle/baselines/word_ppr_deviations.json`

### Spec 010 Phase 3 enablement

- [ ] Hand-derive a `CT_RPr` canonical ordering from `preserved` patterns only
- [ ] Hand-derive a `CT_PPr` canonical ordering from `preserved` patterns only
- [ ] Add corresponding `ChildSequence` entries to `src/openxml_audit/word/compat.py` with comments citing both ECMA-376 sections and the oracle baseline JSON
- [ ] Add unit tests + integration tests mirroring the Phase 1/2 pattern in `tests/test_word_compat_ordering.py`
- [ ] Re-run the mining tool against TokenMoulds with the new constraints active; assert zero false positives

### Documentation

- [ ] Update `docs/word_compat/README.md` validation status table with the empirically-derived canonical orderings
- [ ] Update CHANGELOG with Phase 3 details
- [ ] Issue #3 follow-up comment with the Phase 3 ship status and links

## Phase 4 (Optional): Visual capture

Conditional on a real driver — only if XML-only roundtripping demonstrably misses a finding that screenshots would catch. Not committed by this spec.

- [ ] Wire in TokenMoulds-style screenshot capture as an additional `RoundtripResult` field
- [ ] Add a visual-diff helper for page-by-page comparison
- [ ] Document developer-machine permission requirements (Screen Recording specifically)

## Cross-Phase Hygiene

- [ ] Each phase: re-check ruff and mypy. The `tools/oracle/` tree is allowed to use a more permissive lint config (longer lines for AppleScript embedding, etc.) — document the per-file overrides in `pyproject.toml`
- [ ] Each phase: every committed oracle JSON includes the Word version string from the run that produced it; spec 010 references the version when citing oracle data
- [ ] Each phase: README updates land with the code, not in a separate commit; oracle data lands with its consumer (spec 010 update)
