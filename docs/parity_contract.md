# Parity Contract

This document defines the comparison contract and CI gate behavior used by parity workflows.

## Scope

- Applies to OOXML parity calibration and PR parity-drift gating.
- Baseline target: `Open XML SDK v3.4.1`.
- Baseline artifacts:
  - `data/corpus/sdk_seed/manifest.json`
  - `data/corpus/parity_baseline/v3.4.1/parity_snapshot.json`
  - `data/corpus/parity_baseline/v3.4.1/waivers.json`

## Snapshot Schema

Current snapshots are produced by `scripts/corpus/run_parity_snapshot.py` and include:

- `checks_total`
- `checks_matched`
- `checks_mismatched`
- `checks_missing_files`
- `match_rate_percent`
- `strict`
- `mismatch_families[]` with normalized `description` and `count`

## Comparison Contract

`scripts/corpus/compare_to_baseline.py` compares a current snapshot against baseline and enforces:

- mismatch growth (`checks_mismatched` delta)
- new mismatch-family count
- match-rate drop
- missing-file checks
- check-count drop (unless explicitly allowed)
- strict-mode drift (if required)

Default policy is strict no-drift:

- `--max-mismatch-growth 0`
- `--max-new-families 0`
- `--max-match-rate-drop 0.0`
- `--max-missing-files 0`
- `--require-same-strict`
- disallow `checks_total` drop

## PR Gate Behavior

Workflow: `.github/workflows/parity-gate.yml`

- Runs snapshot with current code.
- Compares to baseline via `compare_to_baseline.py`.
- Uploads reports and step summary.
- Fails PR on policy violations.

Corpus source for gate:

- Local `data/corpus/sdk_seed/files/` if present.
- Otherwise downloads archive from repository variable `PARITY_CORPUS_ARCHIVE_URL`.

Archive requirement:

- Compressed `.tar.zst` that extracts a top-level `files/` directory.

## Waiver Process

Waivers are declared in `data/corpus/parity_baseline/v3.4.1/waivers.json`.

Required fields per waiver:

- `kind`: one of
  - `new_mismatch_family`
  - `mismatch_growth`
  - `match_rate_drop`
  - `missing_files`
  - `check_total_drop`
  - `strict_mode`
- `owner`: accountable owner/team
- `reason`: short rationale
- `expires`: ISO date `YYYY-MM-DD`
- `target`: required for `new_mismatch_family` (family description)

Rules:

- Expired waivers are ignored.
- Invalid waiver entries are ignored and surfaced as warnings in compare output.
- Waivers are temporary by design; renew only with explicit rationale.

Example:

```json
{
  "waivers": [
    {
      "kind": "new_mismatch_family",
      "target": "file_not_found",
      "owner": "openxml-audit-maintainers",
      "reason": "Temporary corpus mirror outage",
      "expires": "2026-04-15"
    }
  ]
}
```

## SDK Usage Policy

- SDK validation remains available for calibration and deep investigations.
- Daily development and PR CI must not depend on local SDK installation.
