<!-- /autoplan restore point: /Users/ynse/.gstack/projects/BramAlkema-openxml-audit/spec-012-parity-gate-recovery-autoplan-restore-20260429-000449.md -->
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

1. **First concrete step: bisect.** Reproduce the snapshot at `5c40485` (Phase 1 HEAD, last green) and `0f8a986` (Phase 2 HEAD, first red), diff the outputs. This is the cheap definitive test for hypotheses (A) vs (B) vs (C) — five minutes settles which surprising families are actually new at Phase 2 vs preexisting. All snapshot outputs land under `docs/parity_recovery_2026-04/transcripts/{short-sha}-snapshot.json` (committed) for reproducibility.
   - Download `openxml-parity-corpus-v3.4.1.tar.zst` from the release URL into `/tmp`.
   - Run `python scripts/corpus/run_parity_snapshot.py --manifest data/corpus/sdk_seed/manifest.json --files-root /tmp/.../files --output docs/parity_recovery_2026-04/transcripts/5c40485-snapshot.json` against Phase 1 HEAD.
   - Run again against `0f8a986` and `f1eeee0` (current `main` HEAD), writing each to its own `{short-sha}-snapshot.json`.
   - Diff the snapshots; confirm or refute that families 1, 2, 4, 5 first appear at `0f8a986`.
2. If Phase 1 already shows families 1, 2, 4, 5: hypothesis (A) is wrong; the regression predates Phase 2. Bisect further back to find the actual introduction commit.
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
- **Pre-push validation:** run `python scripts/corpus/compare_to_baseline.py --baseline data/corpus/parity_baseline/v3.4.1/parity_snapshot.json --current docs/parity_recovery_2026-04/transcripts/post-fix-snapshot.json --waivers data/corpus/parity_baseline/v3.4.1/waivers.json --output /tmp/parity_compare.json --summary /tmp/parity_compare.md --max-mismatch-growth 0 --max-new-families 0 --max-match-rate-drop 0.0 --max-missing-files 0` locally. Confirm exit 0 and no `waiver_warnings` before pushing. This catches typos in `waivers.json` before the gate sees them.
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
5. The recovery PR includes local reproduction transcripts at `docs/parity_recovery_2026-04/transcripts/{short-sha}-snapshot.json` for at minimum: `5c40485` (last green), `0f8a986` (first red), and `main` post-fix (verified green) — so a future maintainer can reproduce the bisect and the recovery delta.
6. Every (B)-classified family in the recovery doc cites the specific transcript file and finding(s) that justify the re-baseline. No (B) classification without an evidence pointer.
7. Spec 010 Phase 2 (`specs/010-word-compat-element-ordering.md`) is updated with a back-link to the recovery doc once classifications are complete, so the next person tracing the parity gate failure finds the resolution from either direction.

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

---

# /autoplan Review

Generated: 2026-04-29 | Branch: `spec-012-parity-gate-recovery` | Commit: `8de51db`
Mode: SELECTIVE EXPANSION (autoplan default for SDK/process iteration on existing system)
UI scope: none → Phase 2 (design review) skipped.

## CEO Review (Phase 1)

### 0A. Premise Challenge

**Stated premise (the spec):** "We need to investigate before waiving/re-baselining because waiver-without-investigation is a parity-contract anti-pattern."

**Steel-manned counter:** *Just re-baseline. The validator is the source of truth; the SDK baseline is a reference. If our validator now reports 5 more families on Document.docx than the SDK does, that's likely improved coverage, not regression. Re-baselining in 5 minutes ships the gate green and unblocks Spec 010 Phase 3.*

**Why the counter loses (verdict):** The four "surprising" families have a property the steel-man ignores: the diff between last-green and first-red is exclusively `src/openxml_audit/word/compat.py` and `compat.py` only emits `Sem_SemanticError`. There is no plausible code path by which Phase 2's diff added `Sch_UndeclaredAttribute` findings. So either the validator's output changed without a code change (configuration drift, harness drift, corpus drift) or the families are pre-existing and the gate just started flagging them. Re-baselining accepts that ambiguity as silently correct. That's the failure mode `docs/parity_contract.md` was written to prevent: "Waivers are temporary by design; renew only with explicit rationale." Re-baselining is the same forfeit at higher altitude.

**Refined premise:** "Investigate enough to *classify* each family as fix/waive/re-baseline. The bisect (Phase 1 step 1, ~5 minutes) is the cheap test. If bisect shows all four surprising families predate Phase 2, the answer is probably 'baseline drift, re-baseline with explicit rationale.' If bisect shows they appear at Phase 2, the answer is 'find the indirect side effect and fix it.' Either way, classification first, action second."

**Outcome we actually want:** A green parity gate on `main` that future maintainers can trust. The intermediate question — fix vs waive vs re-baseline — is downstream of "do we know what changed."

### 0B. Existing Code Leverage

| Sub-problem | Existing code that solves it |
|---|---|
| Run our validator on the corpus and emit a snapshot | `scripts/corpus/run_parity_snapshot.py` |
| Compare snapshot to baseline + enforce gate | `scripts/corpus/compare_to_baseline.py` |
| Re-baseline workflow | `.github/workflows/calibrate-parity.yml` |
| Waiver schema and validation | `docs/parity_contract.md` §Waiver Process; `compare_to_baseline.py` reads `waivers.json` |
| Family-key normalization | `src/openxml_audit/parity_normalization.py` |
| Download corpus archive | `vars.PARITY_CORPUS_ARCHIVE_URL` (release asset) |

The recovery spec writes **zero new code paths**. It is entirely an investigation-and-classification exercise that uses existing tooling. This is the right shape for a parity-recovery operation: any new tooling we'd add (e.g., a "drift triage runbook") is premature with a sample size of one.

### 0C. Dream State Mapping

```
  CURRENT STATE                THIS PLAN                        12-MONTH IDEAL
  ----------------------       ------------------------         --------------------------
  Gate red on main             Gate green on main               Gate has never been red on
  4 commits under              with classified                  main for >24h. Drift events
  failed gate.                 fix/waive/re-baseline            auto-trigger a recovery
  No reproduction              actions and a committed          spec template + bisect
  transcripts.                 reproduction transcript.         script. Waivers age out
  Waivers list empty.          1-3 expiring waivers with        and prompt review.
  Spec 010 Phase 3 +           specific reasons. Recovery
  Spec 011 blocked.            doc as future template.
```

This plan moves us toward the 12-month ideal *without* trying to build the ideal in one shot. Specifically: a "drift triage runbook" is appealing but explicitly out of scope (we have one data point — a runbook from one event is just this PR with extra steps).

### 0C-bis. Implementation Alternatives

```
APPROACH A: Investigate-first (the spec)
  Summary: Bisect + classify each family + targeted fix/waive/re-baseline.
  Effort:  M (human ~6h / CC ~30min)
  Risk:    Low — classification rules out silent regression
  Pros:    - Classification produces audit trail
           - Honors parity contract (reasons required)
           - Each family handled per its actual cause
           - Surprising families get root-cause writeup
  Cons:    - Slower than blind re-baseline
           - Bisect could be inconclusive (mitigated: doc says continue further back)
  Reuses:  All existing parity tooling

APPROACH B: Re-baseline now, investigate later
  Summary: Trigger Calibrate parity, accept new families, file follow-up TODO.
  Effort:  S (human ~1h / CC ~5min)
  Risk:    HIGH — bakes any genuine regression into the new normal
  Pros:    - Fastest unblock for downstream work
           - One workflow run + one PR
  Cons:    - Silently accepts unknown changes
           - Violates parity contract spirit
           - "Investigate later" historically means never
           - If a real regression exists, it's now invisible
  Reuses:  Calibrate parity workflow

APPROACH C: Add waivers blanket-ly, no re-baseline, no investigation
  Summary: Add waivers for all 5 new families with reason="inherited from spec 010 phase 2".
  Effort:  S (human ~30min / CC ~10min)
  Risk:    HIGH — same blindness as B, but with the additional cost of expiring waivers compounding
  Pros:    - Doesn't touch the baseline file
  Cons:    - Vague reasons forbidden by contract
           - Defers the problem 6 months
           - Stacks future renewal burden
  Reuses:  Waivers system
```

**RECOMMENDATION:** APPROACH A. The 5-minute bisect (Phase 1 step 1) is the cheap definitive test. Approach B and C trade investigation cost for opacity, and the contract was specifically written to refuse that trade. Engineering preference: explicit over clever; well-tested over fast.

### 0D. Mode-Specific Analysis (SELECTIVE EXPANSION)

**Complexity check:** The spec touches 0 source files and adds 1 spec file (152 lines, now 162 after corrections). Far under the 8-file/2-class smell threshold. No simplification needed.

**Minimum-set check:** The minimum to ship value is exactly what's in the spec. Already minimal.

**Expansion candidates considered:**

1. *"Drift triage runbook"* — generalize the recovery into a reusable template under `docs/runbooks/parity_drift.md`. **Auto-decided: defer.** Sample size of one. Per Edge case paranoia (design 16): premature templates create false certainty. Revisit after the next drift event.
2. *"Auto-trigger recovery spec template on red gate"* — wire a workflow that opens a draft spec PR when the gate goes red. **Auto-decided: defer to TODOS.md.** Real value but builds a meta-system that needs maintenance. Premature given current cadence.
3. *"Waiver expiry dashboard"* — emit a workflow summary listing all waivers and their expiry dates. **Auto-decided: skip.** With 1-3 waivers ever, the JSON file is the dashboard.
4. *"Add CHANGELOG entry"* — required by acceptance criteria already (Phase 4). Not an expansion.
5. *"Cross-reference Spec 010 Phase 2 with the eventual classification"* — i.e., when we discover the root cause of the surprising families, link back from spec 010. **Auto-decided: include.** Free, prevents the same investigation re-running. Add to acceptance criteria.

**Action:** added expansion #5 to acceptance criteria (see Required Outputs below).

### 0E. Temporal Interrogation

| Hour | What the implementer hits | Decision needed now |
|---|---|---|
| 1 (foundations) | Download corpus archive, where to extract, how to invoke run_parity_snapshot.py | The spec already names the script + flags. ✓ |
| 2-3 (core logic) | Bisect: which commits, in what order, where to write transcripts | Spec names `5c40485`, `0f8a986`, `f1eeee0` and the transcript path. ✓ |
| 4-5 (integration) | What if local repro diverges from CI? | Risks section names this. Falls back to `gh run download`. ✓ |
| 6+ (polish/tests) | Will the recovery PR itself trigger the gate (bootstrapping concern)? | **GAP — surface below.** |

**Bootstrapping concern surfaced:** the recovery PR will run the parity gate. If we add waivers + re-baseline + fixes, the gate evaluates each commit on the branch. If gate config (e.g., waiver expiry parsing) is fragile, the recovery itself could fail mid-way. **Resolution:** spec already says verify gate at end of Phase 4; add an explicit "test waivers parse correctly *before* pushing the recovery PR" step. Edit below.

### 0F. Mode Selection

Confirmed: **SELECTIVE EXPANSION**, with one expansion accepted (#5 above), three deferred. Approach A from 0C-bis.

---

## Required Outputs

### NOT in scope (with rationale)

- **Refactor the parity comparator (`compare_to_baseline.py`)** — works as designed; the failure is upstream of it.
- **Loosen the strict-no-drift gate policy** — defeats the purpose; would mask future regressions.
- **Drift triage runbook** — sample size of one; premature template.
- **Auto-recovery workflow** — meta-system maintenance burden, premature.
- **SDK reference bump (v3.4.1 → v3.5.x)** — orthogonal; own spec.
- **Spec 010 Phase 3 (CT_PPr/CT_RPr)** — blocked on this; design lives in Spec 010.
- **Perf regressions** — only if accidentally introduced by recovery work.

### What already exists (re-stated)

See 0B above. All recovery work uses existing tooling — no new abstractions.

### Dream state delta

This plan moves us from "gate red, no triage record" to "gate green with a classified, reproducible record of what happened and why." It does NOT build the long-term auto-triage system — that's deferred until we have a second drift event to learn from.

### Error & Rescue Registry (recovery process failure modes)

| Step | What can go wrong | How we handle it | What if rescue also fails |
|---|---|---|---|
| Phase 1.1 (download corpus archive) | URL 404, network timeout, bad zstd | Retry once; fall back to `gh run download` of the latest gate run artifact | Manual: ask user to re-publish the release asset; halt spec |
| Phase 1.1 (run snapshot at `5c40485`) | Old commit's deps don't install on current python (e.g., pinned version drift) | `pip install -e .` against a clean venv per checkout | If install fails: skip bisect, classify all surprising families as "needs investigation in CI" and use `gh run download` artifacts from those runs instead |
| Phase 1.3 (find actual finding text) | Snapshot output doesn't contain the unfilled `<value>` — only the family description | Re-emit snapshot with raw (unnormalized) errors using `parity_normalization.normalize_error_tuple`'s pre-templating output | If pre-template output not available: add a one-line debug print to `run_parity_snapshot.py` for the recovery PR only, then revert |
| Phase 3 (waiver entry parses) | Typo in `waivers.json` schema | Run `compare_to_baseline.py` locally before push; check the `waiver_warnings` field in output | If still fails: the gate will surface the warning on push; revise from gate output |
| Phase 3 (re-baseline action) | Calibrate workflow produces unexpected result | Diff calibrate output against current baseline before committing | If diff is large: stop, escalate to user, do not auto-replace baseline |
| Phase 4 (recovery PR triggers gate) | Gate fails on the PR itself due to mid-recovery state | Land each phase in its own commit; gate runs per-push and produces clear feedback | If gate is fundamentally broken: revert to last green commit on the branch, regroup |

### Failure Modes Registry

| Failure | Severity | How we'd notice | Mitigation in plan |
|---|---|---|---|
| Bisect inconclusive (surprising families predate Phase 2) | MED | snapshots at `5c40485` already contain families 1, 2, 4, 5 | Plan explicitly says "bisect further back" — covered |
| Local repro diverges from CI | MED | snapshot output shape differs | Falls back to CI artifacts — covered |
| Recovery PR introduces new mismatch families during fix iteration | LOW | gate red on recovery PR | Per-commit gate runs surface immediately |
| Waiver renewal cliff (6 months) | LOW | waivers expire and gate fails again | Each waiver's `reason` must be specific enough for cheap renewal review — covered |
| User signs off on (B) re-baseline without reading transcripts | MED (process risk) | classification doc says (B) but no transcript link | **GAP — fix below**: add acceptance criterion that any (B)-classified family links to the transcript line(s) supporting the classification |

### Cross-Phase Themes

(Will be filled after Eng phase. Anchor: the bisect-first thread runs through CEO 0A, Eng test diagram, and acceptance criteria. If both phases independently flag bisect-first, that's a high-confidence signal.)

### CEO Dual Voices — Independent Reviews

**CLAUDE SUBAGENT (CEO — strategic independence)**: Read the plan with no prior context.

Major findings (verbatim severities):

1. **HIGH** — *Plan answers the wrong question.* Mission per CLAUDE.md is "will this file open in target apps?" yet the plan treats SDK baseline as the contract. With Spec 010/011 deliberately shipping divergence, SDK is no longer gold standard but a peer with a known incomplete model. Strict-no-drift gate against a peer you've decided is wrong is incoherent. Fix: reframe as "redefine the parity gate's role." Right gates: (a) self-parity vs our previous output, (b) Word/PowerPoint/Excel roundtrip oracle. SDK becomes a third *informational* signal.
2. **MEDIUM** — *"Investigate before waiver" applied as ritual.* Diff is one file, that file emits one finding kind, three surprising families are at structurally unrelated XPaths. Dominant prior is baseline drift, not a validator regression. Decision rule should be: if bisect at `5c40485` already shows surprising families → classify all four as (B) and re-baseline. Caps the work at ~30 min for the likely case.
3. **HIGH** — *No (D) considered: split the gate.* Self-parity + SDK-informational + Word-oracle. The real strategic move; doing the recovery without it guarantees recurrence.
4. **HIGH** — *6-month regret foreseeable.* By Oct 2026: 5–15 expiring waivers, each "Word disagrees with SDK." Waiver list becomes the validator's actual spec, expressed negatively. Fix: don't waive Word-divergence on the SDK gate; record in a `Word divergence registry` with no expiry and positive rationale.
5. **CRITICAL (process)** — *Gate red, main kept moving.* Either gate is required and we halt commits, or it isn't. Fix branch protection or demote gate to advisory in this same PR.
6. **HIGH** — *Recurrence is structural.* Spec 010 Phase 3 + Spec 011 = future drift events. Build minimal `scripts/corpus/classify_drift.py` now (~2 hours).
7. **MEDIUM** — *v3.5 not orthogonal.* 20-min check: does v3.5 already emit some surprising families? If yes, classification settles instantly.

**Bottom line:** Land smaller Spec 012 — (a) demote SDK gate to advisory, (b) re-baseline under that policy, (c) open Spec 013 to design self-parity + oracle gates.

---

**CODEX SAYS (CEO — strategy challenge)**: Independent run on the same plan + repo context. Convergent findings:

- **Wrong sovereign.** "Gate is the contract" assumes SDK is truth; product mission is app survival; specs 010/011 explicitly reject SDK behavior as truth.
- **10x reframe = "separate oracles".** SDK parity only on SDK-equivalent findings; app-compat findings tagged separately and excluded from SDK comparison; Word roundtrip oracle = strategic source of truth for Word-survival claims.
- **Governance failure, not incident.** Four pushes landed → gate is not a gate. Fix branch protection or stop calling it a contract.
- **SDK corpus has too much roadmap power.** Four mismatches all hinge on `TestFiles/Document.docx`. One SDK fixture freezes work that may matter more to users. Not market-weighted quality.
- **"Out of scope: change strict-no-drift" is the dismissal that buries the actual fix.** Six months: every app-survival improvement either needs waivers or breaks CI.
- **Severity model missing.** Plan assumes all drift has equal strategic value. Unknown schema drift may matter; expected Word-rejection matters more; SDK-only disagreement may be noise.
- **Missing alternative:** change `run_parity_snapshot.py` to filter app-compat findings by source so they don't enter the SDK comparison at all. Avoids the false "regression vs waiver" framing.
- **Bulky reproduction transcripts → ceremony.** Optimize for repeatable command, not committed forensic debris.
- **Competitive risk underplayed.** "Matches SDK" = commodity. "Predicts Word/PowerPoint/Excel survival" = moat. Plan defends the commodity benchmark.
- **Spec 011 blocking dependency is backwards.** Oracle should unblock the parity decision, not wait behind it.

---

### CEO Consensus Table

```
═══════════════════════════════════════════════════════════════════════════
  Dimension                                   Claude  Codex  Consensus
  ──────────────────────────────────────────  ──────  ─────  ─────────────
  1. Premises valid?                          NO      NO     CONFIRMED-NO
  2. Right problem to solve?                  NO      NO     CONFIRMED-NO
     (both flag: "split the gate" is the
      real problem; this plan is symptom-fix)
  3. Scope calibration correct?               NO      NO     CONFIRMED-NO
     (both: "out of scope: change policy"
      is dismissing the actual fix)
  4. Alternatives sufficiently explored?      NO      NO     CONFIRMED-NO
     (missing (D): split-gate / source-tag)
  5. Competitive/market risks covered?        NO      NO     CONFIRMED-NO
     (both: validator's moat is app survival,
      not SDK parity; plan defends commodity)
  6. 6-month trajectory sound?                NO      NO     CONFIRMED-NO
     (both: waiver pile-up + recurrence
      structural, not anomalous)
═══════════════════════════════════════════════════════════════════════════
```

**0/6 confirmed-yes. 6/6 confirmed-no.** This is unusually emphatic. Both voices, independently, with different reasoning chains, conclude the plan is structurally misframed — not flawed in detail.

This crosses the autoplan threshold for **USER CHALLENGE** (both models recommend changing the user's stated direction). Surfaced at the Final Approval Gate, not auto-decided.

---

## Engineering Review (Phase 3)

### Step 0: Scope Challenge with Code Reading

I read `scripts/corpus/run_parity_snapshot.py`, `scripts/corpus/compare_to_baseline.py`, `src/openxml_audit/parity_normalization.py`, `src/openxml_audit/word/compat.py`, `data/corpus/parity_baseline/v3.4.1/parity_snapshot.json`, `data/corpus/parity_baseline/v3.4.1/waivers.json`, and `docs/parity_contract.md`. Findings flow through this entire section.

**Scope is fine in size** (1 spec file, no source code). But scope is wrong in *kind* — see CEO User Challenge above. If the plan goes forward as written, the Eng-side issues below need to land first, regardless of CEO-level reframing.

### Architecture (no new components → text instead of diagram)

This plan is *procedural*, not architectural. There are no new components, no new data flows, no new state machines. The only "architecture" is the existing parity-gate pipeline:

```
  CORPUS ARCHIVE ──▶ run_parity_snapshot.py ──▶ parity_current.json
                                                       │
                              parity_snapshot.json     │
                              (frozen baseline)         ▼
                                       │       compare_to_baseline.py
                                       └────▶ ─────────┬─────────────▶ EXIT 0 / EXIT 1
                                              waivers.json (filter)
```

The plan reuses every component as-is. The architecture-review value is "what does the plan assume about each component that may not be true?" — that's where every Eng finding lives. No diagram is novel here; the existing pipeline is well-documented in `docs/parity_contract.md`.

### Test Diagram (mapping recovery work to verification)

| Recovery work | What changed | What test catches breakage | Exists? |
|---|---|---|---|
| Phase 1.1: bisect at `5c40485`, `0f8a986`, `f1eeee0` | New diagnostic procedure | Re-running the bisect should give identical output (deterministic) | NO — relies on local reproduction; no automated regression |
| Phase 3 (W): waiver entries added | `waivers.json` content | Comparator must accept the new entries with no `waiver_warnings` AND the gate must actually pass with them present | NO — `waiver_warnings` are non-fatal; pre-push step is human inspection |
| Phase 3 (B): re-baseline | `parity_snapshot.json` replaced | Future snapshots compared against new baseline; any divergence flagged | YES (the gate itself) |
| Phase 3 (F): fix in code | Source change | Existing unit tests + parity gate | YES |
| Phase 4 verification | Gate green on PR | Gate workflow run | YES |

**Critical gaps:**
- No regression test that asserts the *waivers.json schema* parses cleanly (a typo merges silently as a warning).
- No regression test that asserts the *gate would fail* without the waiver/fix being added (positive control). Without it, dead waivers accumulate forever.
- No test that asserts the recovery's transcripts contain the data they need to (raw error descriptions are not in the snapshot output today; see codex finding #4).

### Eng Dual Voices — Independent Reviews

**CLAUDE SUBAGENT (eng — independent review)**: Read the spec + key implementation files. Major findings:

- **CRITICAL C1** — `pip install -e .` against an old commit may install the *wrong source*. `run_parity_snapshot.py` does `sys.path.insert(0, str(ROOT / "src"))` so it imports from the working tree regardless of pip-install state. Plan must say "git checkout `<sha>` of the entire tree, not just rebuild the venv."
- **CRITICAL C2** — Bootstrapping unresolved. If Phase 3's first waiver lands before its corresponding fix, the gate fails on the intermediate commit. Plan needs squash-merge or temporarily-relaxed gate-on-PR for the recovery branch.
- **HIGH H1** — `family_key` normalization can produce different keys at different commits. Bisect at `5c40485`/`0f8a986` won't catch a normalization-shape change that occurred between baseline-extraction and `5c40485`. Must also include the SHA at which `parity_snapshot.json` was last regenerated.
- **HIGH H2** — `normalize_description` templates out the actual attribute name. Plan's "find the actual finding text" can't be done from committed snapshots. Need a `--debug-raw-descriptions` flag added *as part of this spec* (not deferred).
- **HIGH H3** — `mismatch_families` is capped at top 50, `mismatch_examples` at 200. Must re-run with a much larger cap or verify the ceiling isn't biting.
- **HIGH H4** — `xml.iter()` order in `compat.py:270` depends on lxml parse order; combined with the 20-tuple-per-version slice in `run_parity_snapshot.py:310`, this can spuriously surface or hide families.
- **MEDIUM M1** — Committed transcript JSONs = ceremony. Repeatable shell command beats committed forensic debris.
- **MEDIUM M2** — No regression test for waiver schema. Belongs in `tests/parity/test_waivers_schema.py`.
- **MEDIUM M3** — Re-baseline keeps no forensic history. Should version under `data/corpus/parity_baseline/v3.4.1/snapshots/{date}.json` for forensic diff.

Verdict: "Procedure is *directionally* sound but operationally unsafe."

---

**CODEX SAYS (eng — architecture challenge)**: Independent run against the same code. Convergent + extended findings:

1. **HIGH** — *Waiver path cannot make the gate pass.* `new_mismatch_family` waivers only reduce the *unwaived new-family count* (one of three thresholds). The gate also enforces `mismatch_growth` (delta in `checks_mismatched`) and `match_rate_drop`. Waiving the sectPr family alone won't drop the count of mismatched *checks* (`Document.docx` × 4 versions = 4 mismatched checks even after waiving). The plan is written as if waivers fix all three thresholds; they don't.
2. **HIGH** — *Waiver target wording is wrong.* Spec says `target=family description`; comparator keys by `family_key` (pipe-delimited). Human-readable description won't match. Must use the exact normalized key from `parity_normalization.normalize_error_tuple`.
3. **HIGH** — *"Waiver schema validation" is not validation.* Invalid waivers surface only as `waiver_warnings` (non-fatal). Pre-push step relies on a human noticing them — fragile.
4. **HIGH** — *Transcripts are not forensic enough.* `descriptions_by_version` is collected but never written to JSON. `mismatch_examples` omits raw text. Step 1.3 cannot be done from committed transcripts as planned.
5. **MEDIUM** — *Bisect does not isolate the variable.* Snapshot output is a function of code + manifest + corpus + Python/lxml + dep resolution. "5 min settles A/B/C" is false without recording dep-lock + manifest hash + corpus archive URL/hash + runner image per checkout.
6. **MEDIUM** — *Calibration and gate inputs diverge.* Calibration builds a runtime manifest via `extract_sdk_expectations.py`; parity gate uses checked-in `data/corpus/sdk_seed/manifest.json`. Replacing only `parity_snapshot.json` from calibration can baseline a different check universe than CI will compare. **Re-baselining as the spec describes is broken.**
7. **MEDIUM** — *Family counts are a lossy artifact, not root cause.* Families collected only for mismatched checks, first 20 tuples per version. "New family" may reflect truncation/order changes, not new validator behavior. (Same root as Claude's H3+H4.)

Verdict: "Plan likely produces either non-applying waivers, a still-red pre-push comparison, or an audit trail that cannot explain the raw findings it claims to preserve."

---

### Eng Consensus Table

```
═══════════════════════════════════════════════════════════════════════════
  Dimension                                   Claude  Codex  Consensus
  ──────────────────────────────────────────  ──────  ─────  ─────────────
  1. Architecture sound?                      N/A     N/A    N/A (no new arch)
  2. Test coverage sufficient?                NO      NO     CONFIRMED-NO
     (no waiver-schema test, no bisect repro
      regression, no transcript completeness)
  3. Performance risks addressed?             N/A     N/A    N/A
  4. Security threats covered?                N/A     N/A    N/A
  5. Error paths handled?                     PARTIAL PARTIAL CONFIRMED-PARTIAL
     (Error & Rescue Registry exists but H2/
      codex#4 reveal it can't recover raw text)
  6. Procedure correctness?                   NO      NO     CONFIRMED-NO
     (C1: sys.path; C2: bootstrap; codex#1:
      waiver thresholds; codex#2: target key;
      codex#6: re-baseline broken)
═══════════════════════════════════════════════════════════════════════════
```

**0/4 applicable confirmed-yes. 3/4 confirmed-no, 1/4 confirmed-partial.** The Eng review confirms that even setting aside the CEO-level reframe, the plan as written is operationally broken — the bisect is unsound (cap bias, key-shape drift, sys.path), waivers don't fix all three thresholds, transcripts are not forensic enough, and re-baselining via the calibrate workflow doesn't produce a same-shape file as the gate expects.

---

## Cross-Phase Themes

When themes appear independently in BOTH CEO and Eng dual voices, that's the highest-confidence signal autoplan produces. Two themes did:

**Theme 1: "The plan is symptom-fix on the wrong target."**
- CEO Claude: "Plan answers the wrong question. Mission is app survival, not SDK parity. The right gates are self-parity + Word oracle."
- CEO Codex: "Wrong sovereign. SDK parity becomes informational; oracle becomes truth."
- Eng Claude: "Procedure is *directionally* sound but operationally unsafe — and even if fixed, doesn't address that the SDK gate is the wrong gate."
- Eng Codex: "As written, this plan likely produces non-applying waivers… plus the meta-criticism (split the gate; SDK isn't sovereign) is real but orthogonal."

→ Three independent voices say the gate is wrong; the fourth says even within the gate's frame, the plan doesn't work.

**Theme 2: "The recovery procedure has hidden technical landmines that 5-minute claims paper over."**
- Eng Claude: C1 (sys.path), H1 (family_key drift), H3 (top-50 cap)
- Eng Codex: #1 (waiver thresholds), #2 (target key wording), #5 (bisect provenance), #6 (calibration/gate manifest divergence)
- CEO Claude: "Investigate-first applied as ritual… budget claims 6h; the rule above caps it at 30min" — *but only if the rule applies, which the Eng findings show it can't with the data we have*.

→ The "5-minute bisect" framing was wrong twice over: not 5 minutes (lots of provenance gaps); and even with the data, the conclusion isn't reliable.

**No cross-phase themes found** in: architecture (no new arch), security (no new attack surface), performance (no impact).

---

## Decision Audit Trail

| # | Phase | Decision | Classification | Principle | Rationale | Rejected |
|---|-------|----------|---------------|-----------|-----------|----------|
| 1 | Phase 0 | Set mode = SELECTIVE EXPANSION | Mechanical | Default for "feature enhancement on existing system" per skill rule | Fits the recovery-style work | EXPANSION (no greenfield), HOLD (room for a couple expansions), REDUCTION (already minimal) |
| 2 | CEO 0C-bis | Recommend Approach A (investigate-first) | Mechanical (per plan as written) | P1 (completeness) + P5 (explicit) | Plan's own rationale: contract requires reasons | But see User Challenge below — Eng review shows Approach A is operationally broken; new Approach D (split-gate) emerged from dual voices |
| 3 | CEO 0D #1 | Defer "drift triage runbook" | Mechanical | P3 (pragmatic) | Sample size of one | But see Theme 2 — codex says runbook should land NOW since we have a category, not an incident |
| 4 | CEO 0D #2 | Defer "auto-recovery workflow" to TODOS.md | Mechanical | P3 (pragmatic) | Meta-system burden | — |
| 5 | CEO 0D #3 | Skip "waiver expiry dashboard" | Mechanical | P3 (pragmatic) | JSON file is the dashboard | — |
| 6 | CEO 0D #5 | Include "back-link from spec 010 to recovery doc" | Mechanical | P1 (completeness) | Free, prevents future re-investigation | — |
| 7 | CEO 0E | Add bootstrapping check before pushing recovery PR | Mechanical | P5 (explicit) | Prevents mid-recovery gate failure | — |
| 8 | CEO 0E | Add acceptance criterion: (B)-classified families cite transcript evidence | Mechanical | P1 + P5 | Keeps re-baseline honest | — |
| 9 | Eng test diagram | Note 3 missing tests (waiver schema, bisect regression, transcript completeness) | Taste decision | P1 (completeness) vs P3 (pragmatic) | Adding all 3 tests doubles spec scope; minimum is the waiver-schema test | Surfaced at gate |
| 10 | Cross-phase | Both phases independently flag "split the gate" | **USER CHALLENGE** | N/A — never auto-decided | Both models, different reasoning chains, recommend changing user's stated direction | Surfaced at gate |
| 11 | Cross-phase | Both phases independently flag "5-minute bisect is wrong" | **USER CHALLENGE** | N/A — never auto-decided | Confirms recovery procedure has hidden landmines beyond "investigate vs waive" | Surfaced at gate |

