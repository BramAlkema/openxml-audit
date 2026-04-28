# Spec: Validator-Output Sovereign Gates (self-parity + oracle)

## Status

Proposed (April 29, 2026). Stub opened by Spec 012 Phase 3 to capture the design space; no implementation in this spec yet.

## Problem

Spec 012 demoted the SDK parity gate to advisory because the .NET Open XML SDK is no longer the right sovereign for this validator's output. Specs 010 and 011 deliberately ship divergence from the SDK because the SDK accepts files Word rejects. As a result, we currently have **no blocking gate** on validator-output regressions — only the perf budget remains blocking.

The two `/autoplan` review phases on Spec 012 (CEO + Eng) independently identified that the right replacement is a pair of gates:

1. **Self-parity gate** — keyed on *our* validator's last green output, not the SDK's. Catches real regressions deterministically: any change in our findings vs the previous accepted state needs to be intentional and reviewed.
2. **Oracle gate** — keyed on `Spec 011`'s Word/PowerPoint/Excel roundtrip oracle results. Catches the actual product question: does this file survive the target app? Sovereign for app-survival claims.

Together these replace the SDK-parity-as-blocking model. The advisory SDK comparison stays as a third informational signal so SDK divergence remains visible (useful when bumping SDK ref, when investigating cross-validator differences, etc.).

## Why This Matters

- The current configuration (perf-only blocking) means a regression in validator output that doesn't move the perf needle ships silently.
- Spec 010 Phase 3 (CT_PPr/CT_RPr empirical canonical) and future Spec 011 oracle findings will all introduce intentional output changes. Without a self-parity gate, the human resolver has no canonical "yes I meant this" step. Without an oracle gate, we have no enforcement that intentional changes actually improve app-survival.
- The validator's competitive moat is app-survival prediction, not SDK parity. The blocking gate should match the moat.

## Normative References

- `specs/012-parity-gate-recovery.md` — context for the demotion that created this gap; full `/autoplan` review preserved there.
- `docs/parity_recovery_2026-04.md` — implementation note for the Spec 012 demotion and the manifest's new mixed-semantics status.
- `docs/parity_contract.md` — current "Gate Roles" section; will need an update when this spec lands.
- `specs/011-word-roundtrip-oracle.md` — the oracle work whose output this spec consumes for the oracle gate.
- `.github/workflows/parity-gate.yml` — current workflow; new gates may live here, or in a sibling workflow.

## Approach Sketches (not yet committed)

These are starting points for design, not decisions.

### Self-parity gate

**Core question:** what is "our last green output"? Three forms to evaluate:

| Form | Pros | Cons |
|---|---|---|
| Bytewise snapshot of validator output | Strictest; catches any change including formatting | Brittle; reformatting, ordering, timestamp differences all flag |
| Family-set (set of normalized `family_key` values) | Robust to count fluctuation; captures what kinds of findings exist | Doesn't catch count changes within a family |
| Error-id-set + count per family | Balance of structure and tolerance | More complex baseline format; needs careful rebaseline UX |

**Recommended starting point:** family-set + per-family count, mirroring the existing `parity_snapshot.json` structure but rooted at our output rather than SDK expectations.

**Integration with existing tooling:** `compare_to_baseline.py` is already shaped right. The change is that the baseline becomes "our last green snapshot" instead of "SDK expectations." `run_parity_snapshot.py` continues to produce snapshots; what differs is that `checks_matched` is computed against our own prior output rather than SDK-extracted expectations.

**Re-baseline UX:** when a self-parity gate fails, the resolver decides: was this an intentional change? If yes, run a `gh workflow run rebaseline-self-parity` workflow that re-emits the baseline, which the resolver commits in the same PR as the validator-output-changing change. The commit message must state what changed and why.

### Oracle gate

**Core question:** how does an oracle finding get from Spec 011's roundtripper into a CI gate?

**Sketch:** the roundtripper produces JSON baselines under `tools/oracle/baselines/word_*.json`. A new CI workflow runs the oracle in a deterministic mode against a curated set of regression files, asserts the recorded results match the committed baselines. Drift fails the build.

**Hard part:** the oracle requires actual Word/PowerPoint/Excel installations. CI may not have them. Likely path: oracle gate runs locally on the reviewer's machine before merge, with the workflow asserting the baselines on disk are consistent (signed by a known reviewer's machine). This is a manual-step-with-CI-verification model rather than fully automated.

### Tagging app-compat findings at emission

**Core question:** how do we filter findings by source so the advisory SDK comparison can exclude expected-Word-only findings?

**Sketch:** each `ValidationError` already has an `error_type` (schema, semantic, relationship, etc.). Add a `source_class`: `sdk_proxy`, `word_app_compat`, `oracle_finding`. The advisory SDK comparison filters out non-`sdk_proxy` findings from its diff. The self-parity gate uses all findings.

This enables a clean separation: the SDK signal is purely about SDK-equivalent findings; app-compat findings live in their own gate(s).

### Severity model tied to customer harm

**Core question (per autoplan codex CEO finding):** all drift is currently treated equally. It shouldn't be — unknown schema drift may matter more than expected-Word-rejection divergence; some findings are silent corruption, some are non-issues.

**Sketch:** add a `severity_class` to each finding: `critical` (file unreadable in target app), `high` (visible corruption), `medium` (cosmetic), `low` (informational). The self-parity gate fails on critical/high regressions; flags medium/low for review.

## Open Questions

1. **Baseline format granularity** — family-set vs error-id-set vs full normalized-tuple-set?
2. **How often to refresh** — every PR, every release, on demand? What's the normal flow when a contributor adds a deliberate finding?
3. **Re-baseline approval flow** — purely git-based (commit the new baseline) or workflow-mediated (button click that publishes a signed baseline)?
4. **Oracle gate viability without CI access to Word/PowerPoint** — accept that oracle gate runs locally + signed, or invest in a Windows/macOS CI runner?
5. **Does this spec own the eventual deletion of the advisory SDK gate?** Or do we keep it forever as a third informational signal? (The autoplan voices were split: codex implied "keep as third signal"; claude implied "delete once self-parity + oracle are real.")
6. **Manifest split** — Spec 012 left `data/corpus/sdk_seed/manifest.json` as mixed-semantics (some entries SDK-extracted, some manually adjusted to self-parity). Should this spec split it cleanly into two files, or formalize the mixed mode?
7. **Severity model — who classifies?** Heuristic from finding text? Per-finding metadata at emission time? Curated overrides file?

## Decisions Needed Before Implementation

- Pick a baseline format (open question 1).
- Decide oracle gate integration strategy (open question 4).
- Decide on the SDK gate's long-term role (open question 5).
- Resolve the manifest split (open question 6).

## Out of Scope (for this spec stub)

- Implementation. This stub captures the design space; the next iteration is a real proposal.
- Changes to the existing parity-gate workflow. Spec 012 already demoted SDK; further changes wait for self-parity + oracle to land.
- Changes to `parity_normalization.py` or other validator internals beyond the source-class tagging sketch.
- Performance budget — orthogonal, already blocking.

## Risks (preview)

- The "self-parity baseline" can become stale if not refreshed regularly. Need a clear cadence.
- Oracle gate without CI Word installations is fundamentally a human-in-the-loop gate. That's the actual constraint of the problem space (the SDK can't tell us what Word does), but it limits how strict the gate can be.
- Mixed-semantics manifest from Spec 012 is technical debt that compounds over time.
- Severity classification can become a bikeshed if not pinned to concrete examples early.
