# Spec: Excel Canonical-Form Validators

## Status

Proposed (April 30, 2026). Phase 1 of a focused validator track that
detects patterns Excel will silently rewrite on save. Triggered by
the v0.7.2 baseline finding that all TokenMoulds-emitted `.xlsx`
files come back from Excel with 10 changed parts, despite passing
schema validation cleanly. 0.7.3 ships the first canonical-form
check; subsequent releases add more.

## Problem

The v0.7.2 oracle baseline showed something the validator's
schema/semantic layer didn't catch: every TokenMoulds-emitted `.xlsx`
file had Excel rewriting 10 parts on save, with no repair dialog
(silent canonicalization). A separate investigation identified the
specific causes:

1. **inlineStr cells without `xl/sharedStrings.xml`** — Excel's
   canonical form is shared-strings-table + `<c t="s"><v>idx</v></c>`
   cells. `<c t="inlineStr"><is>...` cells get migrated on save.
2. **Chart references that Excel materializes into
   `xl/externalLinks/`** — when chart `<c:numRef>`/`<c:strRef>`
   refer to cells whose actual content doesn't typematch the chart's
   cached values.
3. **Attribute / whitespace canonicalization** in
   `[Content_Types].xml`, styles, theme, tables, drawings,
   worksheets — undocumented but stable per-element preferences.

The validator's mission is roundtrip survival in the target app.
"Will Excel rewrite this on save?" is precisely a roundtrip
question, even though the file passes formal schema/semantic
checks. This spec opens the canonical-form track to surface those
predictions before the user ships.

## Why This Matters

- **Direct mission relevance.** The v0.7.2 finding ("Excel
  rewrites every TokenMoulds-emitted workbook") is real signal;
  shipping a check that detects the dominant cause turns
  observation into actionable validator output.
- **Findings come pre-tagged for the future sovereign gate.**
  `SourceClass.EXCEL_APP_COMPAT` (Spec 018) means the self-parity
  gate (Spec 026 0.8.0) can include or filter these as the user
  prefers.
- **Each check has a clear test fixture immediately.** The v0.7.2
  corpus under `data/corpus/tokenmoulds_v0.7.2/excel/` triggers
  the inlineStr pattern; future checks will use the same corpus
  + targeted synthetic fixtures.

## Normative References

- `tools/oracle/baselines/README.md` "v0.7.2 clean run" section
  — the empirical motivation. The 10-changed-parts finding on
  `acme-us.xlsx` is what this spec is responding to.
- `src/openxml_audit/errors.py` `SourceClass.EXCEL_APP_COMPAT`
  (Spec 018) — the tag for these findings.
- `src/openxml_audit/excel/workbook.py` `WorkbookValidator` —
  pattern reference for the new validator's shape.
- `data/corpus/tokenmoulds_v0.7.2/excel/{acme-us,globex-gb}.xlsx`
  — committed positive-case fixtures for `Excel_InlineStrCells`.

## Approach

### Phase 1 — Excel_InlineStrCells (this release, 0.7.3)

1. **`src/openxml_audit/excel/canonical_form.py`** —
   `ExcelCanonicalFormValidator` class. Walks every worksheet
   part, counts cells with `t="inlineStr"`, flags worksheets
   that have any when `xl/sharedStrings.xml` is missing or
   empty.
   - Emits one finding per worksheet (not per cell) with the
     count in the description, so the user knows blast radius.
   - Pre-computes "is the SST populated" once per package so
     the check is O(worksheets) rather than O(worksheets × cells).
   - Tags `SourceClass.EXCEL_APP_COMPAT`, severity WARNING,
     id `Excel_InlineStrCells`. Constructs `ValidationError`
     directly (rather than via `context.add_error`) so the
     part_uri can be set per-worksheet.

2. **Wire into `validator.py`** — add a one-liner call inside
   `_validate_spreadsheet_structure` that runs the canonical-form
   validator after the existing workbook structure validator.

3. **`tests/test_excel_canonical_form.py`** — 7 tests:
   - **Positive**: inlineStr cells without SST → flagged;
     inlineStr cells with empty SST → flagged; v0.7.2 corpus
     fixtures (acme + globex) → both flagged with count 9.
   - **Negative**: no inlineStr cells → no finding; sheet uses
     `<c t="s">` with populated SST → no finding; both inlineStr
     AND populated SST → no finding (locks the scope decision
     so the check has high signal-to-noise).

### Phase 2 — Excel_ChartExternalRefMaterialization (later)

Detect chart `<c:numRef>` / `<c:strRef>` whose target cells
don't typematch the chart's cached values. Excel materializes
these into `xl/externalLinks/` on save. Detection requires:

- Resolving the chart's reference (`Sheet1!$B$2:$B$4`) against
  the actual sheet cells in the package
- Comparing each referenced cell's actual type (inlineStr /
  shared-string-ref / number) against the chart's cache
  declaration (`<c:numRef>` expects numeric)
- Flagging when the sheet content can't satisfy the chart's
  type expectation

This is non-trivial; deferred to its own focused release once
Phase 1 stabilizes the canonical-form track.

### Phase 3 — Excel_NonCanonicalAttributeOrder (much later)

Detection requires a reference dataset of "what Excel's
attribute order is per element kind." Corpus mining job
on a body of Excel-saved .xlsx files. Aspirational; not
load-bearing now.

## Acceptance Criteria (Phase 1 / 0.7.3)

1. `src/openxml_audit/excel/canonical_form.py` exists with
   `ExcelCanonicalFormValidator` exposing a `validate(package,
   context) -> None` entry point.
2. Wired into `validator.py`'s `_validate_spreadsheet_structure`
   after the existing workbook structure validator.
3. Findings carry `id="Excel_InlineStrCells"`,
   `severity=WARNING`, `source_class=EXCEL_APP_COMPAT`,
   per-worksheet `part_uri`.
4. The 7 new tests in `tests/test_excel_canonical_form.py` pass.
5. v0.7.1 self-parity baseline passes against itself with the
   new check active (verified: 0 drift; the SDK seed corpus
   doesn't trigger the pattern).
6. Full test suite passes.
7. CHANGELOG updated. Spec 029 committed.

## Out of Scope (Phase 1)

- Phase 2 / Phase 3 checks (deferred above).
- Word-side or PowerPoint-side canonical-form checks. Likely
  exist; their own future spec.
- Auto-fix tooling (e.g. "rewrite this xlsx with proper
  sharedStrings"). The validator only detects.
- Self-parity baseline refresh. Not needed for this release —
  the new check adds zero drift on the SDK seed corpus.

## Risks

- **The check's scope decision** (only flag when SST is
  missing/empty, not when SST is populated AND inlineStr is
  also present) is a deliberate signal-to-noise choice. Some
  files use both; Excel still rewrites those. We chose not to
  flag them here because the dominant motivating case
  (TokenMoulds emitter) has no SST at all. Mitigation: the
  test that locks this decision is explicit; if a future case
  argues for broader detection, the choice is one boolean flip.
- **False positives on hand-crafted files.** A file deliberately
  using inlineStr (e.g. for streaming-write pipelines that don't
  buffer all strings) gets flagged as a WARNING. Acceptable
  cost; the WARNING level (not ERROR) reflects the "this is
  fine, but Excel will rewrite it" semantic.
- **The `ExcelCanonicalFormValidator.validate()` shape.** It
  takes the package and context directly rather than running
  per-part. That's because the canonical-form judgment is
  package-wide (sees both worksheets AND sharedStrings). Future
  checks should follow the same pattern unless they're truly
  per-part.
