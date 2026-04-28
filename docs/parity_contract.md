# Parity Contract

This document defines the comparison contract and CI gate behavior used by parity workflows.

## Gate Roles (as of Spec 012, April 2026)

The repo's mission per `CLAUDE.md` is "validate OOXML files to determine if they will open successfully in their target apps." Specs 010 and 011 deliberately ship divergence from the .NET Open XML SDK because the SDK accepts files Word rejects. The parity-gate roles reflect that:

| Gate | Role | Status |
|---|---|---|
| **SDK parity comparison** (`compare_to_baseline.py`) | Advisory — trend visibility on validator output drift against SDK v3.4.1 | **NON-BLOCKING** in `.github/workflows/parity-gate.yml` (`continue-on-error: true` on the comparison step). Runs every push/PR; surfaces in workflow summary; does not fail the build. |
| **Perf budget** (`check_perf_budget.py`) | Runtime regression guard | **BLOCKING** — perf-budget violations still fail CI. |
| **Self-parity** (validator output vs our last green output) | Catches our own regressions, keyed on our output rather than SDK's | DEFERRED to Spec 013 (`specs/013-validator-output-sovereign-gates.md`). |
| **Word/PowerPoint/Excel roundtrip oracle** | Sovereign for app-survival claims | DEFERRED to Spec 011 + Spec 013. |

The strict-no-drift *measurement* policy (max 0 mismatch growth, 0 new families, 0pp match-rate drop) is unchanged — the comparison still uses those thresholds when deciding what to surface in the summary. The change is what happens when those thresholds are exceeded: the build no longer fails. SDK divergence is now a trend signal, not a gate.

When the SDK reference is bumped (separate spec), the advisory snapshot must be re-extracted so the trend stays meaningful.

## Scope

- Applies to OOXML parity calibration and PR parity-drift gating.
- Baseline target: `Open XML SDK v3.4.1`.
- Baseline artifacts:
  - `data/corpus/sdk_seed/manifest.json`
  - `data/corpus/parity_baseline/v3.4.1/parity_snapshot.json`
  - `data/corpus/parity_baseline/v3.4.1/perf_budget.json`
  - `data/corpus/parity_baseline/v3.4.1/waivers.json`

## Snapshot Schema

Current snapshots are produced by `scripts/corpus/run_parity_snapshot.py` and include:

- `duration_seconds`
- `validation_runs`
- `checks_collected`
- `checks_total`
- `checks_matched`
- `checks_mismatched`
- `checks_missing_files`
- `match_rate_percent`
- `strict`
- `include_mutation_expectations`
- `skipped_mutation_expectations`
- `mismatch_families[]` with normalized tuple fields:
  - `id`
  - `error_type`
  - `part`
  - `path`
  - `description`
  - `family_key` (`id|error_type|part|path|description`)
  - `count`

## Expectation Scenarios

Expectations extracted from SDK tests are tagged in the manifest with `scenario`:

- `base`: validation of an unmodified corpus file
- `mutation`: validation after test-time document mutations (add/remove/edit operations)

Extractor normalization for parity fidelity:

- Only package-level `OpenXmlValidator.Validate(package)` assertions are mapped to file-level expectations.
- Part/element-level validations and max-error harness flows (for example `MaxNumberOfErrors` tests) are tagged as `mutation`.

Default parity snapshots run against `base` expectations only.

- `scripts/corpus/run_parity_snapshot.py` excludes `scenario == "mutation"` unless `--include-mutation-expectations` is passed.
- Excluded counts are reported as `skipped_mutation_expectations` in the snapshot JSON.

## Comparison Contract

`scripts/corpus/compare_to_baseline.py` compares a current snapshot against baseline and enforces:

- mismatch growth (`checks_mismatched` delta)
- new mismatch-family count
- match-rate drop
- missing-file checks
- check-count drop (unless explicitly allowed)
- strict-mode drift (if required)

Family drift is keyed by `family_key` when present (falls back to `description` for older snapshots).

Default *measurement* policy is strict no-drift (advisory; see "Gate Roles" above for the build-failure semantics):

- `--max-mismatch-growth 0`
- `--max-new-families 0`
- `--max-match-rate-drop 0.0`
- `--max-missing-files 0`
- `--require-same-strict`
- disallow `checks_total` drop

These thresholds determine what the comparator surfaces as drift; they no longer determine whether the build passes. Per Spec 012, the SDK comparison is non-blocking.

## PR Gate Behavior

Workflow: `.github/workflows/parity-gate.yml`

- Runs snapshot with current code.
- Compares to baseline via `compare_to_baseline.py` (advisory; see "Gate Roles").
- Enforces runtime budget via `check_perf_budget.py` (blocking).
- Uploads reports and step summary.
- Fails PR only on perf-budget violations. SDK comparison drift is surfaced in the step summary but does not fail the build.

Corpus source for gate:

- Local `data/corpus/sdk_seed/files/` if present.
- Otherwise downloads archive from repository variable `PARITY_CORPUS_ARCHIVE_URL`.

Archive requirement:

- Compressed `.tar.zst` that extracts a top-level `files/` directory.

## Performance Guard

Performance budget is defined in `data/corpus/parity_baseline/v3.4.1/perf_budget.json`.

`scripts/corpus/check_perf_budget.py` validates:

- max total snapshot duration (`max_duration_seconds`)
- max duration per evaluated check (`max_seconds_per_check`)
- max duration per validator run (`max_seconds_per_validation`)
- minimum evaluated checks (`min_checks_total`)

Default policy is fail-on-regression against these limits in PR CI.

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
