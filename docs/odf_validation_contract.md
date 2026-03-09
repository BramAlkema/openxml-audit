# ODF Validation Contract

This document defines the Phase 3 comparison contract for ODF validation in `openxml-audit`.

## Scope

- Applies to OASIS OpenDocument (`.odt`, `.ods`, `.odp`) reference alignment workflows.
- Covers:
  - pinned sample corpus shape
  - normalized run-report schema from `scripts/odf/run_reference_validators.py`
  - mismatch taxonomy schema from `scripts/odf/compare_reference_results.py`
- Does not define full ODF conformance parity requirements.

## Pinned Corpus

- Corpus manifest: `data/odf/reference_corpus/manifest.json`
- Each sample references a fixture directory under `tests/fixtures/odf/`.
- `run_reference_validators.py` materializes deterministic ODF ZIP files from fixture directories.
- `mimetype` is written first and uncompressed when present.

Sample entry contract:

- `id` (string, stable identifier)
- `profile` (string, e.g. `valid` or `invalid`)
- `fixture_dir` (string, path relative to fixtures root)
- `filename` (string, staged output name)

## Run Report Schema

Primary report: JSON output from `scripts/odf/run_reference_validators.py`.

Top-level fields:

- `generated_at` (ISO timestamp)
- `contract_version` (`odf-reference-v1`)
- `corpus_manifest` (resolved path)
- `fixtures_root` (resolved path)
- `strict` (boolean, Python validator mode)
- `sample_count` (integer)
- `duration_seconds` (float)
- `runners` (object with per-runner status counts + command template metadata)
- `samples` (array of sample records)

Per-sample fields:

- `id`, `profile`, `fixture_dir`, `filename`, `staged_file`
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

Skipped samples are recorded when either run status is not `ok`.

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
```

## Known Limitations

- Reference-tool adapters are command-template based and output parsing is best-effort.
- Message-level parity is not guaranteed because external tool formats differ by version.
- No hard CI dependency on Java validators in default workflows.
