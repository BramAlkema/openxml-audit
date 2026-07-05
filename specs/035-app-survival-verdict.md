# Spec: App-Survival Verdict as the Headline Answer

## Status

Proposed (July 5, 2026). Phase 1: rule-based per-app verdict derived
from existing findings, leading the CLI text output and added to JSON
output. Additive only — the SDK-style error list and XML output are
unchanged.

## Problem

The mission (CLAUDE.md, ADR-001) is "will this file open — and
survive — in its target app?", and every finding already carries the
machinery to answer it: `SourceClass` separates SDK-proxy findings
from app-compat rules (cases the SDK accepts but Word/Excel/PowerPoint
repair or rewrite) and package-integrity findings. But
`openxml-audit file.pptx` still leads with a numbered error list — the
user is left to translate findings into the question they actually
asked.

## Design

### Verdict layer (`openxml_audit/verdict.py`)

`predict(result) -> AppVerdict` maps a `ValidationResult` to one
verdict for the file's target app (resolved from its extension:
Word / Excel / PowerPoint / LibreOffice Writer-Calc-Impress).

Rule ladder, most severe first:

1. any ERROR-severity `PACKAGE_INTEGRITY` finding → **reject-likely**
   — the OPC/ODF package layer is broken; apps cannot be expected to
   open it.
2. any `<App>_APP_COMPAT` finding for the target app, **at any
   severity** → **repair-or-rewrite-likely** — these rules encode
   observed app behavior (repair dialogs, silent canonicalization
   rewrites). Severity does not gate this rung: the Spec 029
   shared-strings rule is warning-severity yet predicts a rewrite the
   v0.7.2 oracle baseline confirmed on every TokenMoulds workbook.
3. any other ERROR-severity finding (`SDK_PROXY`, `ODF_NATIVE`) →
   **at-risk** — schema/semantic violations; the app may open,
   repair, or reject.
4. otherwise → **opens-clean** (warnings noted).

Each verdict states its evidence basis in words and carries the top
findings that drove it. Honesty constraint: these are rule-derived
predictions from the validator's own finding taxonomy, not
statistically calibrated claims. Measured precision/recall against
oracle-ground-truth corpora is future work and a prerequisite for any
percentage claim.

### CLI

- Text output: one verdict headline per file, above the existing
  error list. The list, `Errors:`/`Warnings:` trailers, and XML
  output are byte-compatible with today's format.
- JSON output: each file object gains a `"verdict"` block
  (`app`, `prediction`, `headline`, `evidence`, `basis`).
- Exit codes unchanged.

## Non-Goals

- No probability or percentage claims until verdicts are scored
  against oracle baselines on a real corpus.
- No secondary-app verdicts (Google Workspace, LibreOffice-for-OOXML)
  until findings carry evidence classes for those apps.
- No change to validation behavior, error content, or the parity
  surface.

## Acceptance

- `openxml-audit reference.pptx` leads with
  "Microsoft PowerPoint: expected to open cleanly …" for a clean file.
- A file with Excel app-compat findings leads with
  "Microsoft Excel: expected to repair or rewrite this file …".
- A file with package-integrity findings leads with reject-likely.
- JSON output carries the verdict block; XML output is unchanged.
- All existing CLI tests pass unmodified.
