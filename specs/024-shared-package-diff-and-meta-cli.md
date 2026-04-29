# Spec: Shared Package Diff Module + Meta-CLI

## Status

Proposed (April 29, 2026). Phase 1 shipping in 0.6.8 — extracts the
PPTX-specific per-part diff machinery in `pptx.lab` into a
format-agnostic `openxml_audit.package_diff`, refactors the XLSX /
ODF / Word corpus oracles to use it (replacing hash-only diffs with
canonical-c14n + per-part text diffs), and adds a single dispatcher
at `python -m openxml_audit.oracle <engine> ...`.

## Problem

After 0.6.7 the four roundtrip oracles each had their own diff path:

  - PPTX: `pptx.lab.compare_pptx_packages` — full canonical-c14n +
    per-part text diff + timing-tree change collector. Highest-quality
    signal.
  - ODF: hash-only over four canonical parts. Tells you something
    changed; can't tell you *what*.
  - XLSX: hash-only over a list of canonical parts. Same limitation.
  - Word: hash-only over a list of canonical parts. Same limitation.

The hash-only path is too coarse for Phase 2 work — when Excel "found
a problem" with `winsemius.tokens.xlsx`, the 0.6.6 baseline could only
report `preserved` (canonical bytes matched on the parts we
fingerprinted) or `repaired` (something differed somewhere). It
couldn't tell us *which* attribute changed, what was added, what was
removed. The PPTX oracle didn't have this limitation because
`pptx.lab` shipped the better machinery.

Three friction points the unification fixes:

1. **Per-format duplication.** Each oracle re-implemented the same
   "load parts → compare hashes → record changed list" loop with
   minor variations. New diffs cost N×file edits.
2. **Asymmetric signal quality.** PPTX gave per-part text diffs;
   the others gave hash deltas. The validator's mission is
   roundtrip prediction across all four formats with comparable
   evidence.
3. **CLI sprawl.** `python tools/oracle/word_repair_corpus.py`,
   `python tools/oracle/xlsx_repair_oracle.py`, etc. — four paths
   to remember when one verb would do.

## Why This Matters

- **Phase 2 of Specs 019/020/021/022** needs per-part text diffs to
  categorize repairs (cosmetic XML reflow vs. substantive content
  edit). The shared module is the precondition.
- **Spec 013's eventual sovereign gate** consumes oracle output. If
  the four oracles emit comparably-shaped reports, the gate is one
  consumer, not four.
- **Future fifth-format additions** (e.g., a Visio `.vsdx` oracle,
  a Google-Docs-export oracle, anything else ZIP-of-XML) reuse the
  module instead of re-implementing the same diff path.

## Normative References

- `src/openxml_audit/pptx/lab.py`'s
  `_load_package_parts` / `_canonicalize_xml` /
  `_compare_package_parts` / `_write_part_diff` / `_pretty_part_text`
  / `_sanitize_part_name` — extraction source. The functions are
  unchanged in behavior; only their location is.
- Spec 022's `tools/oracle/baselines/README.md` — the "Phase 1 hash
  diff is too coarse" caveat that motivates this work.
- Spec 011 (Word oracle) and the four corpus walkers
  (`word_repair_corpus`, `xlsx_repair_oracle`, `pptx_repair_oracle`,
  `odf_repair_oracle`) — consumers refactored to use the new module.

## Approach

### Phase 1 — extract + refactor + meta-CLI (this release, 0.6.8)

1. **`src/openxml_audit/package_diff.py`**: new module exporting
   `canonicalize_xml`, `load_package_parts`, `compare_package_parts`,
   `write_part_diff`, `pretty_part_text`, `sanitize_part_name`, and
   the high-level `compare_packages(base_path, head_path,
   output_dir, *, parts_filter=None, max_diff_files=50)`.
   - The format-agnostic seam is the `parts_filter` callable. Pass
     `None` for "any .xml/.rels in the ZIP" (default) or a stricter
     callable to scope to canonical parts of one format (e.g.
     PPTX's `ppt/slides/`/`ppt/slideLayouts/`/etc., ODF's four
     top-level XML parts, XLSX's `xl/worksheets/`/`xl/styles.xml`/
     etc.).
   - Canonicalization uses lxml's c14n with `remove_blank_text=True`
     so trivial whitespace reflow doesn't show as a change. This is
     the property that makes the diff fair against repair-on-save.
   - `pretty_part_text` falls back to raw decoded text for
     malformed XML AND for the `recover=True` returns-None case
     (latent bug fix surfaced by the new test suite).

2. **`src/openxml_audit/pptx/lab.py`**: imports the extracted
   primitives under their original `_-prefixed` names (so existing
   PPTX consumers stay binary-compatible). The
   `compare_pptx_packages` public API is unchanged; the
   timing-tree change collector remains PPTX-specific.

3. **`tools/oracle/{xlsx,odf,word}_repair_*.py`**: each oracle's
   per-part diff path was hash-only. Replaced with
   `compare_packages(...)` calls that emit per-part text diffs
   under `<work_dir>/compare/diffs/`. Observation gains
   `added_parts`, `removed_parts`, `diff_dir` (None unless
   `--keep-artifacts`).

4. **`src/openxml_audit/oracle/__main__.py`**: dispatcher under
   `python -m openxml_audit.oracle`. Subcommands:
   - `word`, `excel` (alias `xlsx`), `pptx` (alias `powerpoint`),
     `odf` — route to the matching corpus oracle.
   - `preflight` — runs `tools/oracle/preflight.py` for all four
     engines.

5. **`tests/test_package_diff.py`**: 11 new tests. Cover the
   canonicalization (whitespace-only diffs collapse), parse-error
   fallbacks, custom `parts_filter`, end-to-end report shape,
   sanitization, and unified-diff output.

### Phase 2 — repair categorization on top of the diff (later)

The per-part text diffs are now collectable; Phase 2 categorizes
each diff:

  - **Cosmetic**: only attribute order, namespace prefix, or blank
    nodes changed (post-c14n these should already be filtered, but
    catch the long tail).
  - **Substantive — additive**: parts added, content extended.
  - **Substantive — destructive**: parts removed, content shortened
    (the validator's high-priority signal — "Excel removed
    something").
  - **Format-specific**: PPTX's timing-tree changes,
    Word's revision-mark changes, etc., layered on top.

The categorizer slots in as a new function consuming the existing
`compare_packages` report; format-specific extensions live in the
per-format `lab` modules.

## Acceptance Criteria (Phase 1 / 0.6.8)

1. `openxml_audit.package_diff` module exists and exports the seven
   public functions.
2. `pptx.lab.compare_pptx_packages` unchanged in observable
   behavior; PPTX consumers (validator tests, lab CLI) still pass.
3. XLSX, ODF, and Word corpus oracles produce per-part diffs under
   `<work_dir>/compare/diffs/` when called with their CLIs and
   `--keep-artifacts`.
4. `RoundtripObservation` for XLSX and ODF gains `added_parts`,
   `removed_parts`, `diff_dir`. Word's corpus walker
   (`word_repair_corpus.py`) gains the same.
5. `python -m openxml_audit.oracle preflight` runs the
   four-engine preflight check.
6. `python -m openxml_audit.oracle word|excel|pptx|odf FILES...`
   dispatches to the right corpus walker.
7. New test suite `tests/test_package_diff.py` passes with 11
   tests covering canonicalization, filter, end-to-end, fallbacks.
8. CHANGELOG updated. Spec 024 committed.

## Out of Scope (Phase 1)

- Repair categorization (Phase 2; needs the diffs Phase 1 ships).
- PPTX timing-tree change collector relocation. It stays in
  `pptx.lab` because it consumes the snapshot output that lives
  there. A future cleanup might lift it into a `pptx.diff_extras`
  module, but that's separate work.
- Word's `tools/oracle/word_window.py` migration to
  `src/openxml_audit/docx/osa.py`. Same Word-oracle scope-creep
  caveat from Spec 023; the matrix-driven Word oracle
  (`word_repair_oracle.py`) depends on the existing layout.
- Auto-discovery of file format in the meta-CLI (i.e.
  `python -m openxml_audit.oracle FILE` without `<engine>`).
  Useful but adds another argument-parsing layer.

## Risks

- **Behavior drift in `compare_pptx_packages`.** Mitigation:
  existing PPTX tests pass; the moved functions retain their
  exact behavior (only their locations changed).
- **Latent `recover=True` returns-None case in
  `pretty_part_text`.** Surfaced by the new tests in this release.
  Now handled with a None check + raw fallback.
- **Path conflicts in stacked diffs.** If two corpus walks emit
  to the same `output_dir`, per-part diffs from the second
  overwrite the first. Mitigation: each oracle's `compare_dir`
  is per-staging-run already; cross-run conflicts only arise if
  the caller passes a shared `output_dir` deliberately (Phase 2
  may add namespacing).
