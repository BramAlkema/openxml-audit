# ODF reference baseline — 2026-07-20

Re-baseline of the ODF reference-calibration gate. Supersedes
`../2026-03-09/`, which was taken over 23 samples on 2026-03-09; the corpus
grew to 126 samples on 2026-03-13 (commit `c4c9c89`, "Reach 100% ODF parity
126/126") and the baseline was never regenerated, so the gate compared 126
current samples against 23 baselined ones under zero-drift thresholds and had
been unpassable since. It failed every scheduled run (2026-04-01, 2026-07-01)
and every manual dispatch.

## What this baseline is

A saved compare report — the same schema the gate produces each run — from CI
run 29751353983 on `main`, after the staged-path canonicalisation fix (PR #8).
The reference tools are ODF Toolkit v0.13.0 and OPF validator v0.20-alpha-2
over the 126-sample corpus in `data/odf/reference_corpus/manifest.json`.

## What blessing it accepts as ground truth

Blessing does **not** assert Python and the reference validators agree. It
freezes the *current* disagreement as the accepted state, so future runs gate
on change from here rather than on the March snapshot. At capture:

| tool | samples | only_python | only_reference |
|---|---:|---:|---:|
| odf_toolkit | 126 | 142 | 339 |
| opf | 126 | 142 | 689 |

- **only_python** (142) — findings the Python validator emits that the
  reference tools do not. These are the candidate false positives. The top
  families are ODF-native structural rules (`content.xml` references a style
  but `styles.xml` is not declared in the manifest; missing `office:body`;
  manifest entry not found in package; frame missing position/size). They read
  as legitimate Python-side rules the reference tools do not implement, but
  **this baseline blesses them without per-family confirmation** — they are
  frozen as accepted, not verified correct.
- **only_reference** (339 / 689) — findings the reference tools emit that
  Python does not, i.e. Python's coverage gaps. Freezing them here means new
  gaps are gated; it does not close the existing ones.

The full per-family breakdown is in `mismatch_triage.md`;
`prev_drift_vs_2026-03-09.json` records the drift from the old baseline that
motivated the re-baseline.

## Follow-up

The 142 only_python families deserve a pass to confirm they are true ODF
conformance findings rather than Python false positives. Blessing was a
deliberate choice to make the gate live again; it is not a substitute for that
triage.
