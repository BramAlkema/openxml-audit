# Spec: Self-Parity Prototype as Advisory CI

## Status

Proposed (April 29, 2026). Phase 1 of Spec 026 — ships the
self-parity snapshot + comparator + advisory CI workflow. 0.8.0
promotes the comparator to the blocking sovereign gate.

## Problem

After 0.7.0, the validator has the prerequisites for self-parity
but no actual self-parity tooling: there's a `family_key`
normalization (`parity_normalization.normalize_error_tuple`),
`SourceClass` tagging on findings, and the SDK
`compare_to_baseline.py` shows the shape a parity comparator
should look like — but no script reads our own validator's output
into a baseline-vs-current diff.

This spec lands the missing pieces as advisory CI. Drift surfaces
in workflow summaries; nothing fails the build yet. That sequencing
is deliberate: it lets the format evolve before 0.8.0's blocking
promotion without breaking anyone.

## Why This Matters

- **The substance of 0.8.0 is here.** 0.8.0 is policy
  (continue-on-error → enforced, branch protection update); the
  code is this release.
- **Format choice (Spec 013 OQ1) gets answered concretely.** The
  proposal — family-set + per-family count, `source_class` tagged
  — gets exercised against the real corpus before it's committed
  to as a contract.
- **Two parallel signals expose redundancy and divergence.**
  After this release ships, every push runs both the SDK advisory
  parity comparison and the self-parity advisory comparison; the
  next few weeks of CI summaries show which signal catches what
  the other misses.

## Normative References

- `specs/026-self-parity-sovereign-gate-roadmap.md` — the umbrella
  this Phase 1 implements.
- `specs/013-validator-output-sovereign-gates.md` — original stub
  with the open questions.
- `src/openxml_audit/parity_normalization.py` — `family_key`
  shape and 5-tuple normalization.
- `src/openxml_audit/errors.py` — `SourceClass` enum (Spec 018).
- `scripts/corpus/compare_to_baseline.py` — SDK comparator;
  threshold-flag UX is mirrored here.
- `.github/workflows/parity-gate.yml` — SDK advisory workflow;
  `self-parity-gate.yml` parallels its shape.

## Approach

### Phase 1 — prototype + advisory CI (this release, 0.7.1)

1. **`scripts/parity/run_self_parity_snapshot.py`** — walks the
   manifest's files, runs the validator at each
   `FileFormat.<version>`, accumulates a `family_inventory`
   keyed by `family_key`. Per-family count, `source_class` tag,
   plus first-seen sample fields (raw `description`, raw `path`,
   `part_uri`) so the templated `<value>` placeholder in the
   normalized key doesn't lose diagnostic detail. Output schema
   `version 1`.

2. **`scripts/parity/compare_self_parity.py`** — diffs current vs
   baseline snapshot. Reports `new_families`, `missing_families`,
   `count_drift`. Three threshold knobs:
   `--max-new-families` / `--max-missing-families` /
   `--max-count-drift-total`. Strict-no-drift defaults (all 0).
   Exit nonzero on threshold violation; renders a markdown summary
   for workflow surfaces.

3. **`data/corpus/self_parity_baseline/v0.7.1/snapshot.json`** —
   the initial baseline captured at this release's `main` HEAD on
   the existing parity corpus (887 files, 18,516 findings, 1,368
   unique family_keys, 228 of which are tagged
   `word_app_compat` from Spec 010 Phase 1+2).

4. **`.github/workflows/self-parity-gate.yml`** — runs on every
   push/PR. Snapshot → compare → upload reports → write summary.
   Compare step has `continue-on-error: true` (advisory).
   Coexists with the existing SDK `parity-gate.yml`.

5. **`tests/test_self_parity.py`** — 10 unit tests covering the
   compare logic (zero drift on identical, new/missing/drift
   detection, threshold flags, absolute-drift accounting,
   markdown rendering), plus a smoke that locks in the committed
   baseline's schema shape against itself.

### Phase 2 — blocking promotion (0.8.0)

Spec 026's 0.8.0 step. Removes `continue-on-error: true`, updates
branch protection, settles SDK gate's long-term role.

## Acceptance Criteria (Phase 1 / 0.7.1)

1. `scripts/parity/run_self_parity_snapshot.py` exists and produces
   schema-version-1 output containing `family_inventory`,
   `by_source_class`, and `total_findings`.
2. `scripts/parity/compare_self_parity.py` exists, supports the
   three threshold flags, exits 0 on no drift and nonzero on
   any.
3. `data/corpus/self_parity_baseline/v0.7.1/snapshot.json` exists
   and the comparator passes when invoked with that file as both
   `--baseline` and `--current`.
4. `.github/workflows/self-parity-gate.yml` runs on push/PR with
   `continue-on-error: true` on the compare step.
5. `tests/test_self_parity.py` passes; full test suite passes.
6. CHANGELOG updated. Spec 027 committed.

## Out of Scope (Phase 1)

- **Blocking promotion** — that's Spec 026 0.8.0.
- **TokenMoulds-API corpus rebaseline** — Spec 026 0.7.2.
- **Repair categorization** — Spec 026 0.7.3.
- **Branch protection update** — only when self-parity flips
  blocking in 0.8.0.
- **Schema migration tooling** — `schema_version: 1` is what
  this release ships; if 0.8.0+ ever changes the shape, schema
  evolution gets its own spec.
- **Multi-baseline comparison** (e.g. compare against the previous
  3 baselines to detect oscillation) — interesting future work,
  not load-bearing now.

## Risks

- **Format evolution.** The "family-set + per-family count" shape
  is a hypothesis. If it turns out a different shape is needed
  (e.g. include path normalization stability checksum, or split
  by file/version axis), 0.7.1's advisory status means we can
  evolve before 0.8.0. Mitigation: schema_version field in the
  output makes future migration explicit.
- **Corpus drift between snapshots.** The comparator assumes the
  same corpus on both sides. If the manifest or files change
  between baseline and current, the diff conflates corpus drift
  with validator drift. Mitigation: 0.7.1 doesn't try to solve
  this; the same caveat applies to the SDK parity gate.
- **`source_class` tag completeness.** If a finding is emitted
  by code that hasn't been swept for source-class tagging
  (Spec 018 noted this as future work), it'll come back as
  `sdk_proxy` by default. The baseline still works; the
  classification just isn't fully accurate. Mitigation: a
  source-class adoption sweep is parallel work.
- **Initial baseline includes 228 `word_app_compat` findings.**
  These are real (Spec 010 Phase 1+2 catches them) but mean the
  baseline has an inherent floor that any change to those
  validators will trip. That's correct behavior — intentional
  validator changes need to refresh the baseline.
