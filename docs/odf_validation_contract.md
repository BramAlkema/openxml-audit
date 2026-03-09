# ODF Validation Contract

This document defines the Phase 4 calibration contract for ODF validation in `openxml-audit`.

## Scope

- Applies to OASIS OpenDocument (`.odt`, `.ods`, `.odp`) reference alignment workflows.
- Covers:
  - pinned sample corpus shape
  - normalized run-report schema from `scripts/odf/run_reference_validators.py`
  - mismatch taxonomy schema from `scripts/odf/compare_reference_results.py`
  - drift-gate policy and waiver model from `scripts/odf/check_reference_drift.py`
- Does not define full ODF conformance parity requirements.

## Pinned Corpus

- Corpus manifest: `data/odf/reference_corpus/manifest.json`
- Each sample references a fixture directory under `tests/fixtures/odf/`.
- `run_reference_validators.py` materializes deterministic ODF ZIP files from fixture directories.
- `mimetype` is written first and uncompressed when present.

Sample entry contract:

- `id` (string, stable identifier)
- `profile` (string, e.g. `valid` or `invalid`)
- `category` (string mismatch family grouping seed, e.g. `package`, `schema`, `semantic`, `security`)
- `fixture_dir` (string, path relative to fixtures root)
- `filename` (string, staged output name)
- `file_format` (string, currently `odf1.2` or `odf1.3`)
- optional metadata (for example `odf_version_marker`)

## Run Report Schema

Primary report: JSON output from `scripts/odf/run_reference_validators.py`.

Top-level fields:

- `generated_at` (ISO timestamp)
- `contract_version` (`odf-reference-v1` for run reports, `odf-reference-v2` for compare reports)
- `corpus_manifest` (resolved path)
- `fixtures_root` (resolved path)
- `strict` (boolean, Python validator mode)
- `sample_count` (integer)
- `duration_seconds` (float)
- `python_issue_categories` (category aggregate for python findings)
- `runners` (object with per-runner status counts + command template metadata)
- `samples` (array of sample records)

Per-sample fields:

- `id`, `profile`, `category`, `fixture_dir`, `filename`, `file_format`, `staged_relpath`
- `runs`:
  - `python` (always attempted)
  - `odf_toolkit` (optional; unavailable if command not configured)
  - `opf` (optional; unavailable if command not configured)

Per-run fields:

- `status`: one of `ok`, `unavailable`, `timeout`, `error`
- `duration_seconds` (when executed)
- `exit_code` (when executed)
- `issues` (normalized issue rows)
- optional diagnostics (`reason`, `stdout_preview`, `stderr_preview`, `command`)
- Note: runtime/bootstrap failures (for example missing Java runtime) are classified as
  `status=unavailable` or `status=error` and do not contribute issue rows.

## Normalization Rules

Python rows:

- Use `openxml_audit.parity_normalization.normalize_error_tuple`.
- Add:
  - `severity` from `ValidationError.severity`
  - `comparison_key = "<severity>|<normalized_description>"`

Reference rows (ODF Toolkit / OPF):

- Parse JSON payloads when possible; otherwise parse text output lines.
- Severity inference:
  - contains `warn` -> `warning`
  - contains `info` -> `info`
  - otherwise -> `error`
- Description normalization uses `normalize_description`.
- Comparison key:
  - `comparison_key = "<severity>|<normalized_description>"`

## Comparison Contract

`scripts/odf/compare_reference_results.py` compares Python vs each reference tool independently.

- Comparison unit: `comparison_key`
- Matching logic: multiset (`Counter`) intersection/difference per sample
- Per-tool outputs:
  - sample counts (`total`, `compared`, `skipped`)
  - issue totals (`python`, `reference`, `matched`, `only_python`, `only_reference`)
  - mismatch families (`only_python`, `only_reference`) sorted by count
    - each row includes `family_group_key` for cross-tool family grouping
  - mismatch categories (`only_python`, `only_reference`) grouped by issue category
- Cross-tool grouping:
  - top-level `cross_tool_families.only_python` and `cross_tool_families.only_reference`
  - grouping key is normalized from `comparison_key` after tool-name/path/file noise reduction
  - grouped rows include `count` and per-tool count breakdown (`tools`)

Skipped samples are recorded when either run status is not `ok`.

## Drift Gate Contract

`scripts/odf/check_reference_drift.py` compares a current compare report to a pinned baseline and enforces
threshold policy.

Policy file:

- `data/odf/reference_baseline/2026-03-09/drift_policy.json`

Default strict policy:

- `max_only_python_growth = 0`
- `max_only_reference_growth = 0`
- `max_new_only_python_families = 0`
- `max_new_only_reference_families = 0`
- `max_compared_sample_drop = 0`
- `max_unavailable_samples = 0`
- `max_timeout_samples = 0`
- `max_error_samples = 0`

Failure conditions:

- mismatch growth beyond threshold
- new mismatch-family count beyond threshold (after waiver filtering)
- compared-sample drop beyond threshold
- unavailable/timeout/error sample counts beyond threshold

Gate output:

- JSON report with per-tool baseline/current deltas and family drift:
  - default `reports/odf/reference_drift.json`
- Markdown summary for CI step summaries:
  - default `reports/odf/reference_drift.md`

## Waiver Model

Waivers are declared in `data/odf/reference_baseline/2026-03-09/waivers.json`.

Required fields:

- `kind`
- `owner`
- `reason`
- `expires` (`YYYY-MM-DD`)

Optional fields:

- `tool` (tool-scoped waiver)
- `target` (required for family-targeted waiver kinds)

Allowed waiver kinds:

- `only_python_growth`
- `only_reference_growth`
- `new_only_python_family` (`target` required)
- `new_only_reference_family` (`target` required)
- `samples_compared_drop`
- `reference_unavailable`
- `reference_timeout`
- `reference_error`

Rules:

- Expired waivers are ignored.
- Invalid waiver entries are ignored and surfaced as warnings.
- Waivers are temporary and owner-accountable by design.

## CI Calibration Workflow

Workflow: `.github/workflows/odf-reference-calibration.yml`

- Triggered on schedule and manual dispatch.
- Executes:
  - `run_reference_validators.py`
  - `compare_reference_results.py`
  - `build_mismatch_triage.py`
  - `check_reference_drift.py`
- Enforces drift policy using:
  - baseline compare report
  - `drift_policy.json`
  - `waivers.json`
- Bootstraps external validators in-workflow via:
  - `scripts/odf/bootstrap_reference_validators.py`
- Dispatch inputs can pin/override validator refs:
  - `odf_toolkit_ref`
  - `opf_ref`

## Reproducibility

To regenerate baseline artifacts:

```bash
python scripts/odf/run_reference_validators.py \
  --corpus-manifest data/odf/reference_corpus/manifest.json \
  --output data/odf/reference_baseline/2026-03-09/reference_runs.json

python scripts/odf/compare_reference_results.py \
  --input data/odf/reference_baseline/2026-03-09/reference_runs.json \
  --output data/odf/reference_baseline/2026-03-09/mismatch_report.json \
  --summary data/odf/reference_baseline/2026-03-09/mismatch_summary.md

python scripts/odf/build_mismatch_triage.py \
  --compare data/odf/reference_baseline/2026-03-09/mismatch_report.json \
  --runs data/odf/reference_baseline/2026-03-09/reference_runs.json \
  --output data/odf/reference_baseline/2026-03-09/mismatch_triage.md

python scripts/odf/check_reference_drift.py \
  --baseline data/odf/reference_baseline/2026-03-09/mismatch_report.json \
  --current data/odf/reference_baseline/2026-03-09/mismatch_report.json \
  --policy data/odf/reference_baseline/2026-03-09/drift_policy.json \
  --waivers data/odf/reference_baseline/2026-03-09/waivers.json \
  --output data/odf/reference_baseline/2026-03-09/drift_report.json \
  --summary data/odf/reference_baseline/2026-03-09/drift_summary.md
```

## Known Limitations

- Reference-tool adapters are command-template based and output parsing is best-effort.
- Message-level parity is not guaranteed because external tool formats differ by version.
- ODF calibration depends on external validator build/install success in CI.

## Reference Runner Troubleshooting

- Ensure command templates resolve to runnable commands in the execution environment.
- For automated setup, run `scripts/odf/bootstrap_reference_validators.py`.
- Ensure Maven is installed (`mvn -version`) or Docker is available.
- To force Docker for both build and validator runtime:
  - `python scripts/odf/bootstrap_reference_validators.py --maven-mode docker --runtime-mode docker`
- If a runner reports `unavailable`, inspect `reason` in `reference_runs.json`.
- Use `stdout_preview` / `stderr_preview` from run reports to verify parser-compatible output.
- Template placeholders:
  - `{file}`, `{file_dir}`, `{file_name}`, `{file_stem}`, `{file_suffix}`
  - if no placeholder is provided, file path is appended.
