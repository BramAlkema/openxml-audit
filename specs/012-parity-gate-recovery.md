# Spec: Parity Gate Recovery After Spec 010 Phase 2

## Status

Proposed (April 29, 2026). Parity gate has been red on `main` since commit `0f8a986` (Spec 010 Phase 2, April 28) — four pushes have landed under a failing gate.

## Problem

The Parity gate (`.github/workflows/parity-gate.yml`) compares our Python validator's findings against a frozen `Open XML SDK v3.4.1` baseline (`data/corpus/parity_baseline/v3.4.1/parity_snapshot.json`). Baseline policy is strict no-drift: max 0 mismatch growth, 0 new families, 0pp match-rate drop.

Since Spec 010 Phase 2 landed, every push to `main` has tripped all three gates:

```
Mismatch growth:           4   (threshold 0)
New mismatch families:     5   (unwaived 5, waived 0)
Match-rate drop:          5.19pp  (threshold 0.00pp)
```

The 4 mismatched checks are all on `TestFiles/Document.docx`:

| Validator version | Expected (SDK) | Actual (us) | Delta |
|---|---|---|---|
| Office2007 (inline_version_count) | 415 | 416 | +1 |
| Office2010 (inline_version_count) | 0 | 1 | +1 |
| Office2013 (inline_version_count) | 1 | 2 | +1 |
| Office2013 (assert_single_validator) | 1 | 2 | +1 |

The 5 new mismatch families are:

| # | Family | Count | Expected vs surprising |
|---|---|---|---|
| 1 | `Sch_UndeclaredAttribute` — `…/sdt/sdtContent/tbl/tr/trPr/cnfStyle` | 12 | **Surprising** |
| 2 | `Sch_UndeclaredAttribute` — `…/sdt/sdtContent/tbl/tblPr/tblLook` | 6 | **Surprising** |
| 3 | `Sem_SemanticError` — sectPr child reordering | 4 | **Expected** (this is what Phase 2 ships) |
| 4 | `Sem_UniqueAttributeValue` — `…/AlternateContent/Fallback/pict/shapetype` | 2 | **Surprising** |
| 5 | `Sch_UndeclaredAttribute` — `…/sdt/sdtContent/tbl/tr/tc/tcPr/cnfStyle` | 1 | **Surprising** |

The "expected" family aligns with what Phase 2 was built to ship: the sectPr canonical-ordering check now flags Document.docx's section properties. Adding a `new_mismatch_family` waiver for that family is the documented contract path (`docs/parity_contract.md` §Waiver Process).

The four "surprising" families are not from Phase 2's diff. The diff between `5c40485` (last green) and `0f8a986` (first red) modifies exactly one source file: `src/openxml_audit/word/compat.py`. That file only emits `Sem_SemanticError` findings via `WordCompatValidator.validate()`. It cannot directly emit `Sch_UndeclaredAttribute` or `Sem_UniqueAttributeValue`. Either:

- **(A)** Phase 2 has an indirect side effect that surfaces dormant schema findings (e.g., it perturbs an element-iteration order that other validators implicitly depend on), or
- **(B)** These findings have been emitted by us all along but the baseline was extracted in a state where they were not present (e.g., the SDK corpus archive at the parity-gate URL differs from what was used to seed the baseline), or
- **(C)** A third commit on the path (build/install env, scripts, etc.) altered how the snapshot script normalizes output.

The spec must determine which of (A)/(B)/(C) is true *before* deciding fix-vs-waive on each family. Waiving a regression silently is a parity-contract anti-pattern.

## Why This Matters

- **The gate is the contract.** A red gate on `main` for 4 commits means every subsequent change ships under unverified parity. The signal is now noise.
- **Waiver-without-investigation is forbidden by the contract.** `docs/parity_contract.md` requires `reason` per waiver and treats waivers as *temporary*. We don't yet have a defensible reason for families 1, 2, 4, 5.
- **Re-baselining without investigation hides regressions.** The baseline is meant to track our validator's *intentional* divergence from the SDK. Bumping it without understanding what changed converts every silent regression into the new normal.
- **Spec 010 Phase 3 + Spec 011 are blocked behind this.** Phase 3 (CT_PPr/CT_RPr empirical canonical) will introduce more deliberate divergence from the SDK and needs a clean gate to land safely. The oracle work in Spec 011 will too.

## Normative References

- `docs/parity_contract.md` — comparison contract, waiver process, perf budget
- `.github/workflows/parity-gate.yml` — strict-no-drift policy enforced on every PR/push
- `.github/workflows/calibrate-parity.yml` — re-baselining workflow (manual + weekly cron)
- `data/corpus/parity_baseline/v3.4.1/` — current baseline (snapshot, waivers, perf budget)
- `scripts/corpus/run_parity_snapshot.py` — produces current snapshot
- `scripts/corpus/compare_to_baseline.py` — comparator + gate
- Spec 010 Phase 2 commit `0f8a986` — the change that broke the gate
- Spec 010 Phase 1 commit `5c40485` — last green parity gate run

## Current Failure Pattern

1. Push to `main` containing `src/**` change.
2. Parity gate downloads corpus archive from `vars.PARITY_CORPUS_ARCHIVE_URL`.
3. `run_parity_snapshot.py` validates the corpus, produces `parity_current.json`.
4. `compare_to_baseline.py` compares against `parity_snapshot.json`, emits 5 new families and 4 mismatched checks.
5. Gate fails.
6. Subsequent pushes inherit the failure unchanged — none of the 4 commits since Phase 2 modify the validator core, so the gate output is identical.

## Approach

A four-phase recovery: **investigate → triage → land fixes/waivers → verify → close**.

### Phase 1 — Reproduce locally and identify the source of each surprising family

Goal: convert "surprising" into "explained" for families 1, 2, 4, 5.

Steps:

1. Reproduce the parity-gate flow locally:
   - Download `openxml-parity-corpus-v3.4.1.tar.zst` from the release URL into `/tmp`.
   - Run `python scripts/corpus/run_parity_snapshot.py --manifest data/corpus/sdk_seed/manifest.json --files-root /tmp/.../files --output /tmp/parity_current.json` against `main` HEAD (commit `f1eeee0`).
   - Confirm we reproduce the 5 new families.
2. Bisect to confirm Phase 2 (`0f8a986`) is the actual transition point:
   - Run the snapshot at `5c40485` (Phase 1 HEAD).
   - Run at `0f8a986` (Phase 2 HEAD).
   - If Phase 1 already shows families 1, 2, 4, 5: hypothesis (A) is wrong; the regression predates Phase 2. Bisect further back.
3. For each surprising family, find the actual finding text in the validator output (the `<value>` in the family description templates out the real attribute name) and identify the validator code path that emits it.
4. Cross-check: did Spec 010 Phase 1 actually pass the gate? Inspect the green run's report artifact (`gh run download 25037832958 -n parity-gate-reports`) to confirm the green snapshot did not contain these families.
5. Decide for each surprising family whether it is:
   - **(F) Real regression** — fix the validator (or whatever caused the change in output)
   - **(W) Intentional new finding** — add a waiver with `reason` documenting why we now diverge from the SDK
   - **(B) Baseline drift** — the baseline was extracted under different conditions; re-baseline is appropriate

### Phase 2 — Triage and document each family

For each of the 5 families, record in `docs/parity_recovery_2026-04.md` (or extend the existing parity contract):

| Field | Content |
|---|---|
| family | Description and path |
| count | Total instances in current snapshot |
| classification | F / W / B (per Phase 1 step 5) |
| root cause | One-paragraph explanation |
| action | "fix in commit X" / "waive with reason Y" / "re-baseline because Z" |
| owner | Who is on the hook |
| follow-up | Any deferred work |

Family 3 (sectPr reordering) is already classified W — Phase 2 was built to ship it. Action: waiver with reason "Spec 010 Phase 2 — intentional divergence from SDK; SDK accepts reordered sectPr children, Word does not."

### Phase 3 — Land fixes and waivers

- For (F) families: land code fixes in atomic commits, each one bisectable. Re-run the snapshot locally between commits to confirm progress.
- For (W) families: extend `data/corpus/parity_baseline/v3.4.1/waivers.json` per the contract format (kind=`new_mismatch_family`, target=family description, owner, reason, expires=YYYY-MM-DD). Default expiry: 6 months out (2026-10-31), forces a review.
- For (B) families: trigger `Calibrate parity` workflow (`gh workflow run calibrate-parity.yml -f sdk_ref=v3.4.1`), download the artifact, replace `parity_snapshot.json` and re-commit. Document the trigger in the same recovery doc.

### Phase 4 — Verify and close

- Re-run `Parity gate` workflow on the recovery PR.
- Confirm: 0 mismatch growth, 0 new unwaived families, 0 match-rate drop.
- If perf budget tripped during investigation, address separately (out of scope here unless we accidentally regressed perf).
- Land the recovery PR. Verify gate is green on `main`.
- Update `CHANGELOG.md` with the recovery summary (what changed, what waivers added, expiries).

## Acceptance Criteria

1. Parity gate passes on `main` HEAD with strict-no-drift policy unchanged (`--max-mismatch-growth 0 --max-new-families 0 --max-match-rate-drop 0.0`).
2. Every waiver in `waivers.json` has an explicit `reason` field that names the spec or commit that justifies the divergence.
3. `docs/parity_recovery_2026-04.md` (or equivalent in-tree note) records the classification (F/W/B) and root cause for all 5 families. No silent classifications.
4. No code regression hides under a waiver: every (F) family has a corresponding fix commit referenced in the recovery doc.
5. The recovery PR includes a local reproduction transcript (snapshot output before/after) committed under `docs/parity_recovery_2026-04/transcripts/` so a future maintainer can reproduce.

## Out of Scope

- **Refactoring the parity comparator itself** (`compare_to_baseline.py`) — keep the contract stable; this spec adapts to it.
- **Changing the strict-no-drift policy** — loosening `--max-new-families` or `--max-match-rate-drop` to make the gate pass would defeat the purpose. If the policy needs tuning, that's a separate spec.
- **Perf-budget regressions** — only address if we accidentally introduce one during recovery; otherwise track separately.
- **SDK reference bump** (e.g., to v3.5.x) — orthogonal; let it be its own spec.
- **Spec 010 Phase 3 (CT_PPr/CT_RPr empirical canonical)** — explicitly blocked on this recovery, but its design and execution belong in spec 010, not here.
- **Generalizing the recovery into a "drift triage runbook"** — appealing but not now; we'll have one data point. Revisit after the next drift event.

## Risks

- **Local reproduction may diverge from CI.** The Parity gate uses `python 3.11` on `ubuntu-24.04` with a specific archive URL. Reproducing on darwin may produce different output. Mitigation: if local repro fails, replicate in a docker container or rely on CI artifact downloads (`gh run download`) for ground truth.
- **Phase 1 step 2 (bisect) might show the regression predates Phase 2.** If the green Phase 1 run masked these findings due to corpus or harness differences, the "surprising" families could be much older. The plan still applies — investigate, classify, act — but the writeup needs to acknowledge the pre-existing condition.
- **Re-baselining (B-classified families) is irreversible without rolling back the snapshot file.** Mitigation: any (B) action requires explicit user sign-off in the recovery PR description, not a silent baseline bump.
- **Waivers expire.** A 6-month expiry forces a review but also creates a future cliff. Mitigation: each waiver's `reason` must be specific enough that the renewer can re-evaluate cheaply. Vague reasons ("inherited from spec 010") are forbidden.
