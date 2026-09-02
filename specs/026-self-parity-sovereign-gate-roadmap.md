# Spec: Self-Parity Sovereign Gate — Roadmap to 0.8.0

## Status

Implemented in 0.8.0 (September 2, 2026). Self-parity is the blocking
Forgejo status check on canonical `main`; the .NET SDK v3.5.1 comparison
remains advisory as an external calibration signal.

## Problem

After 0.7.0, the validator has rich app-survival tooling (four
roundtrip oracles, per-part diffs, auto-dismiss, source-class
tagging) and two committed baseline runs as evidence — but **no
blocking gate** on validator-output regressions. The SDK parity
gate is advisory (since Spec 012); only the perf budget is
blocking.

This is the structural gap Spec 012's `/autoplan` review identified
as the highest-leverage move:

> "The right gates are (a) self-parity (regression detection on our
> own output) and (b) Word/PowerPoint/Excel oracle truth (Spec 011).
> The SDK becomes a third signal, informational, not a blocker."
>
> — Spec 012 autoplan CEO consensus, 0/6 confirmed-yes against the
> stated direction.

Spec 013 opened the design space; that stub is now ready to
implement. 0.8.0 is the right release to land it because:

- All four oracles ship and emit standardized observations
  (per-part diffs in 0.6.8, source-class tagging in 0.6.2).
- The validator's output format is stable enough to baseline
  against (path-indexing fix in 0.6.1; family_key normalization
  in `parity_normalization.py`).
- The CHANGELOG narrative for 0.7.0 explicitly listed "Self-parity
  sovereign gate (Spec 013)" as a deferred 0.7.x/0.8.x candidate.
- Branch protection on `main` already disables the SDK gate as
  required (since Spec 012); the slot for the new sovereign gate
  is open.

## Why This Matters

- **Closes the structural gap from Spec 012.** The autoplan
  voices wanted "split the gate" — Spec 012 demoted the SDK gate
  to advisory; 0.8.0 promotes self-parity to the blocking gate.
  Together they complete the reframe.
- **Restores a real regression-detection contract** without
  re-introducing the SDK as gold standard. Self-parity catches the
  same class of regression (anything that changes our validator's
  output) but is keyed on our own output, not the SDK's.
- **Keeps the validator's "no .NET required" promise faithful.**
  Self-parity is pure Python at runtime; the .NET runtime tool from
  Spec 013 OQ8 stays as a developer-side investigation aid only.
- **Sets up the oracle-driven gate** as the ultimate truth for
  app-survival claims (Spec 011 + later release).

## Normative References

- `specs/012-parity-gate-recovery.md` — context for the demotion
  that created this gap; full `/autoplan` review preserved there.
- `specs/013-validator-output-sovereign-gates.md` — the stub
  this implements. Open Question 1 (baseline format) and Open
  Question 5 (SDK gate's long-term role) are answered by this
  release sequence.
- `docs/parity_contract.md` — gets the "Gate Roles" section
  rewritten in 0.8.0 to mark self-parity as sovereign.
- 0.7.0 CHANGELOG narrative — the "deferred to future releases"
  list this spec works through.

## Approach — the four-step path

### 0.7.1 — Self-parity prototype as advisory CI (Phase 1)

Lay the new gate's primitives without changing existing policy.

1. **`scripts/parity/run_self_parity_snapshot.py`** — emits a
   baseline-shaped JSON keyed by `family_key` (the existing 5-tuple
   from `parity_normalization.normalize_error_tuple`) plus
   per-family count. Each entry tagged with the
   `SourceClass` from Spec 018. The structural shape mirrors the
   existing SDK `parity_snapshot.json` so the comparator can be
   thin.
2. **`scripts/parity/compare_self_parity.py`** — diffs current vs
   baseline. Same threshold flags as `compare_to_baseline.py`
   (`--max-mismatch-growth`, `--max-new-families`,
   `--max-match-rate-drop`). Strict-no-drift policy by default.
3. **`data/corpus/self_parity_baseline/v0.7.1/snapshot.json`** —
   the initial baseline captured at this version's `main` HEAD on
   the existing SDK seed corpus.
4. **`.github/workflows/self-parity-gate.yml`** — runs the snapshot
   + compare on every push/PR. **Advisory at this stage**
   (`continue-on-error: true`); the workflow runs but doesn't fail
   the build. SDK `parity-gate.yml` keeps running unchanged.
5. **Documentation:** `docs/parity_contract.md` gets a small note
   that self-parity is now an active second signal alongside SDK
   parity, both advisory pending Spec 013's blocking-promotion
   release.

Tests: unit tests for the new snapshot script (synthetic
`ValidationError` inputs → expected snapshot shape) and the
comparator (synthetic baseline + current → expected verdict).

### 0.7.2 — TokenMoulds-API corpus helper + clean re-baseline

Address the corpus-curation gap flagged in 0.6.9. Fix the corpus
problem so the baselines have defensible provenance before the
sovereign-gate flip.

1. **`tools/oracle/build_corpus.py`** — drives TokenMoulds'
   Python API to emit concrete `.docx` / `.xlsx` / `.pptx` / `.odt`
   documents (not templates). Uses TokenMoulds'
   `tokenmoulds.emitters.{word,excel,powerpoint,odf}` directly,
   bypassing the CLI's font-metrics-cache resolution issue. Output
   under `data/corpus/tokenmoulds_v<rev>/`.
2. **Re-run all four oracles** on this clean corpus via
   `openxml-audit-oracle <engine>`. Commit JSON observations
   under `tools/oracle/baselines/<format>/0.7.2-clean.json`.
3. **Update `tools/oracle/baselines/README.md`** — the "headline
   narrative" becomes "this is what TokenMoulds *actually* emits;
   here's how each Office app behaves on it" instead of the
   mixed-quality claim from 0.6.9.
4. **Re-run the self-parity baseline** (from 0.7.1) on this
   refreshed corpus, replacing
   `data/corpus/self_parity_baseline/v0.7.1/snapshot.json` with a
   `v0.7.2/snapshot.json`.

Tests: smoke tests for `build_corpus.py` (TokenMoulds-API
availability, output file shape).

### 0.7.3 — Word oracle in-package consolidation + repair categorization

Operational hygiene before the sovereign-gate flip.

1. **Fold `tools/oracle/word_window.py` + `word_roundtrip.py`
   into `src/openxml_audit/docx/osa.py`** — the in-package osa
   layer already has `launch_word`, `open_document`, `save_document`,
   `close_document`. This release adds `close_document_saving`,
   `find_repair_dialog_text`, `dismiss_repair_dialog`,
   `dismiss_any_leftover_modal`, `click_dialog_button`,
   `REPAIR_DIALOG_PATTERNS`, `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`
   matching the PPTX and XLSX modules.
2. **Update `tools/oracle/word_repair_corpus.py`** to consume
   `docx.osa` instead of `tools/oracle/word_window.py`.
3. **Update `tools/oracle/word_repair_oracle.py`** (the matrix
   runner) to use `docx.osa` — careful, this is Spec 010/011's
   primary oracle.
4. **Delete `tools/oracle/word_window.py` + `word_roundtrip.py`**
   once nothing references them.
5. **Add `categorize_repair(observation)` to
   `openxml_audit.package_diff`** — walks per-part diffs and
   flags `cosmetic` (attribute reorder, whitespace) vs
   `substantive` (element added/removed, content edited) vs
   `unknown` (couldn't determine).
6. **Re-run the four oracles** on the 0.7.2 clean corpus and
   commit categorized baselines under
   `tools/oracle/baselines/<format>/0.7.3-categorized.json`.

Tests: cover the categorizer with synthetic diffs + known
classifications.

### 0.8.0 — Self-parity becomes blocking; SDK retired or
permanent-advisory

Policy flip. The substance was 0.7.1; this release commits to it.

1. **Promote `.forgejo/workflows/self-parity-gate.yml`** by removing
   `continue-on-error: true` from the comparison step. The
   workflow now fails the build on regression.
2. **Forgejo branch protection on `main`**: require the `self-parity`
   status check in the canonical repository.
3. **Decide SDK gate's long-term role** (Spec 013 OQ5):
   - **Option K (keep advisory)**: `parity-gate.yml` continues
     running as an informational third signal alongside the
     blocking self-parity gate and the eventual oracle gate.
     Codex CEO voice in Spec 012 advocated this.
   - **Option R (retire)**: delete `parity-gate.yml`,
     `calibrate-parity.yml`, `extract_sdk_expectations.py`, and
     the `expectations[]` lists in `manifest.json`. Claude CEO
     voice in Spec 012 advocated this once self-parity + oracle
     were real.
   The 0.8.0 release CHANGELOG records the choice with rationale.
   Decision: **Option K** for the 0.8 release line so the trend signal
   stays available; **revisit at 0.9.0** once the oracle gate's design
   is settled.
4. **Rewrite `docs/parity_contract.md` "Gate Roles" section**:
   self-parity is sovereign, SDK is informational (or retired,
   per the choice above), oracle observations are the
   app-survival truth.
5. **`CHANGELOG.md` 0.8.0 entry**: cutover narrative covering
   the four-step path, the SDK gate decision, and what's still
   deferred (oracle gate sovereignty, repair categorization
   refinement).

## Acceptance Criteria (each step)

| Release | Criterion |
|---|---|
| 0.7.1 | Self-parity advisory workflow exists, runs on every push/PR, never fails the build. SDK parity-gate continues unchanged. |
| 0.7.2 | All four oracle baselines re-shipped against TokenMoulds-API output. README narrative defensible. |
| 0.7.3 | `tools/oracle/word_window.py` deleted; matrix oracle still passes; repair categorization shipped. |
| 0.8.0 | Self-parity gate is required on `main`. Branch protection updated. SDK gate decision documented. |

## Out of Scope (this roadmap)

- **Oracle-driven gate sovereignty.** Spec 013 named it as the
  third pillar (alongside self-parity and SDK informational); a
  future spec works through how to wire oracle observations into
  a CI gate (the chicken-and-egg problem: oracles run on dev
  machines with Office apps installed; CI doesn't have those).
- **Localization of dialog patterns.** Non-English Office builds
  have different button labels — already noted in Spec 023.
- **Visio (`.vsdx`) or other fifth-format oracle.** Same shape as
  the existing four, but its own work.
- **Parity-snapshot schema versioning.** The current
  `schema_version` field is `1` everywhere; if 0.8.0+ ever
  changes the shape, schema migration deserves its own spec.

## Risks

- **Self-parity baseline format choice (Spec 013 OQ1) might be
  wrong.** The proposal here is "family-set + per-family count" —
  same shape as SDK parity, just keyed against our output instead
  of SDK expectations. If a different format (error-id-set,
  full-tuple-set) proves better in practice, a future release can
  evolve it. Mitigation: 0.7.1 ships the format as advisory; it
  can change before 0.8.0's promotion without breaking anyone.
- **TokenMoulds-API corpus helper (0.7.2) may hit the same
  font-metrics-cache resolution problem the CLI did.** Mitigation:
  call TokenMoulds' Python API directly rather than the CLI;
  use the existing `examples/build_acme_*.py` as the pattern
  reference (those work).
- **Word oracle migration (0.7.3) is the most invasive change.**
  The matrix-driven Word oracle (Spec 010/011) is a lot of code
  that consumes `tools/oracle/word_window.py`. Mitigation:
  migrate by changing imports only; keep the function signatures
  identical; verify the existing constraint baselines still
  pass after the migration.
- **The SDK gate decision at 0.8.0 is genuinely contested.** The
  autoplan voices split. Mitigation: Option K (keep advisory) is
  reversible; Option R (retire) commits to deletion. Default to K
  for one release, revisit at 0.9.0 with more data.

## Sequencing notes

- Each 0.7.x patch can ship independently; the sequence is the
  recommended order, not a hard dependency.
- 0.7.1 (self-parity prototype) is the highest-priority because
  it's the substance of 0.8.0; the others (0.7.2, 0.7.3) can
  slip without blocking the cutover.
- If the corpus helper (0.7.2) is harder than expected, ship the
  0.8.0 cutover anyway with a "the baselines are still the 0.6.9
  noisy ones; corpus refresh deferred" caveat. Self-parity
  sovereignty doesn't depend on a clean corpus to be useful — it
  detects regressions against whatever the current baseline is.
