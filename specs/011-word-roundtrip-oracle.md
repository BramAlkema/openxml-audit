# Spec: Word Roundtrip Oracle for App-Compat Findings

## Status

Proposed (April 28, 2026)

## Problem

Spec 010 is blocked behind a proxy-vs-oracle gap. The validator now flags WordprocessingML property-element reorderings that *probably* trigger Word's "unreadable content" repair dialog, but the certainty rests on:

- ECMA-376 XSD ordering, which the .NET Open XML SDK runtime contradicts
- SDK schema metadata `Children` lists, which preserve declarative order but say nothing about Word's runtime tolerance
- A real-world DOCX corpus (TokenMoulds) where 100% of `tblPr`/`tcPr`/`sectPr` and 99.78% of `pPr` observations align with the proxy, but 1.4% of `rPr` observations (334/23,754) deviate from the proxy in a way that's almost certainly Word-tolerated

Each of those is a proxy. None tells us what Word actually repairs.

The same applies to any future Word-app-compat finding beyond ordering: the only authoritative signal is "Word's behavior on this input." Today we have no way to elicit that signal programmatically.

A sibling project — [TokenMoulds](https://github.com/BramAlkema/TokenMoulds) — already has the equivalent for PowerPoint, in [`tools/visual/powerpoint_capture.py`](https://github.com/BramAlkema/TokenMoulds/blob/main/tools/visual/powerpoint_capture.py) (1,399 LOC) and [`tools/visual/pptx_window.py`](https://github.com/BramAlkema/TokenMoulds/blob/main/tools/visual/pptx_window.py) (255 LOC): an osascript-driven harness that opens a `.pptx` in real PowerPoint, drives slideshow/window state, and captures evidence of PowerPoint's behavior. That pattern is the obvious template for a Word equivalent.

## Why This Matters

- Phase 3 of spec 010 (`CT_PPr`/`CT_RPr`) cannot ship a defensible canonical ordering without ground-truth: we have to know which of the corpus deviations Word actually accepts vs repairs.
- Phase 4 of spec 010 (`CT_TrPr` refinement) is currently waiting on Shaun's corpus from issue #3. A roundtrip oracle removes the dependency: we can construct the trPr pairwise-swap matrix ourselves and resolve each cell empirically.
- Future Word-compat findings — element ordering is just the first one; namespace handling, MC alternate content, font substitution, and others all need the same kind of "what does Word do with this?" signal.
- The validator's stated mission is roundtrip survival in the target app. We've been inferring "would this survive Word?" from proxies. The roundtripper turns inference into observation.

## Normative References

- Spec 010 — `specs/010-word-compat-element-ordering.md`. The immediate consumer.
- TokenMoulds PPTX visual harness — `~/projects/TokenMoulds/tools/visual/{powerpoint_capture,pptx_window}.py`. Pattern reference for Word; not a copy target (the API surfaces differ).
- Apple osascript / AppleScript / JXA documentation — for window discovery, UI scripting, and Word app dispatch.
- Microsoft Word for Mac (Microsoft 365) AppleScript dictionary — `osascript -e 'tell application "Microsoft Word" to ...'`. Word exposes a richer scripting surface than UI-only PowerPoint operations require.
- Issue #3 — the original repair-dialog report this is meant to make verifiable.

## Current Failure Pattern

1. Spec 010 ships constraints derived from the SDK schema metadata's `Children` list.
2. The TokenMoulds corpus shows the proxy holds for 4/6 surveyed property types but fails for `rPr` (and is borderline for `pPr`).
3. With no oracle, the team faces a forced choice per type: either ship the proxy (risking false positives if Word is more permissive than the SDK) or hand-derive a canonical from corpus dominance (risking false negatives if the dominant patterns happen to be the ones Word repairs silently).
4. The validator's user-facing severity ladder muddles the picture: WARNING is appropriate for proxy findings but understates real Word-rejected files; ERROR over-promises on proxy findings. Without an oracle, we can't tell which.
5. Issue #3's gold standard — "files that triggered Word's repair dialog with their offending property tree preserved" — exists exactly nowhere in any data store we control.

## Decisions

1. **Tier 1 first, Tier 2 deferred.** Build the XML-only roundtripper (open in Word → Save through Word → return the post-Word file) before anything visual. Visual capture (screenshots, page-by-page rendering) is genuinely useful but not required for spec 010's needs and adds substantial complexity.
2. **Live under `tools/oracle/`, not `src/openxml_audit/`.** This is developer-machine infrastructure. It requires macOS, a Microsoft Word license, UI scripting permissions, and Screen Recording permissions. Putting it under `tools/` makes the boundary clear: not shipped with the wheel, not part of the importable validator, not exercisable in CI.
3. **macOS only for v1.** Document the Linux (LibreOffice as a different-but-not-equivalent oracle) and Windows (PowerShell-driven Word automation) paths as future work, but don't block on them.
4. **Repair dialog → "Yes, accept repair."** When Word's "unreadable content" dialog appears, dismiss as Yes. The post-Word file then represents what Word made of the input, which is the oracle. Record the dialog appearance separately (binary signal: was repair triggered).
5. **Output format is JSON, one file per oracle run.** Each scenario in the run gets an entry with input fragment, post-Word fragment, repair-dialog flag, and verdict (`preserved` / `repaired`). Schema documented in this spec.
6. **Reusable engine + thin spec-010 glue.** The roundtripper itself takes any DOCX and returns a post-Word DOCX. The spec-010-specific concern (generate property-ordering scenarios, diff per parent type) is a small driver layer on top — keeps the engine reusable for future spec-011-style oracle work.
7. **No CI integration.** Tests that depend on Word are marked with a `requires_word_app` pytest marker and excluded from the default suite. The oracle is run by hand, results are committed as JSON.

## Scope

### In Scope

- An osascript/AppleScript-driven harness that opens a DOCX in Microsoft Word for Mac, saves the file through Word, detects repair dialogs, and returns the saved file.
- A driver that generates spec-010 ordering scenarios (initial set: every pairwise swap inside `CT_TrPr`; the corpus-mined `rPr` deviation patterns from TokenMoulds), runs them through the roundtripper, and produces a JSON oracle report.
- Documentation in `tools/oracle/README.md` covering setup (Word license, permissions, first-run prompts), invocation, and how to consume the oracle output to refine spec 010 constraints.
- Pytest marker `requires_word_app` plus a small set of unit-tested helpers (XML diff, scenario generation, report serialization) that don't need Word.

### Out of Scope

- Visual capture (screenshots, slideshow equivalents, page-by-page rendering). Defer to a Tier 2 spec or fold in once Tier 1 is proven.
- LibreOffice-based roundtripping. LibreOffice and Word disagree about repair behavior; LibreOffice can't substitute for Word as the oracle for "Word will repair this."
- Linux or Windows ports. Out of scope for v1 — flag as future-work bullet.
- PowerPoint or Excel equivalents. Each app would need its own spec; this one is Word-only.
- Office 365 web, Google Docs, or any non-desktop-Word target. Different oracles, different specs.
- Refinement of spec 010 constraints based on oracle output. That's spec 010 Phase 3/4 work, not this spec — this spec ships the oracle and the data; spec 010 consumes it.

## Goals

1. Roundtrip a DOCX through Word and capture the post-Save XML, with a one-line API.
2. Detect Word's repair dialog and surface it as a structured signal, not as an exception.
3. Produce a deterministic JSON oracle report for spec 010 scenarios — committable so the team can read and review without rerunning the harness.
4. Document developer-machine setup so a fresh contributor can run the oracle within an hour, including the macOS permission prompts.
5. Stay testable in CI for the parts that don't need Word — keep the ratio of "depends on Word" to "depends on logic" small enough that contributors can debug the logic without a Word license.
6. Establish the `tools/oracle/` pattern: future Word-app-compat findings should plug in as new scenario drivers without re-engineering the roundtripper.

## Design

### Module Layout

```
tools/oracle/
  README.md                     # setup, permissions, first-run, troubleshooting
  __init__.py
  word_window.py                # osascript helpers: launch, open, close, dialog inspection
  word_roundtrip.py             # engine: roundtrip(docx) -> Path
  word_repair_oracle.py         # spec-010 driver: scenarios -> roundtrip -> JSON report
  scenarios/
    __init__.py
    property_ordering.py        # scenario generators per CONSTRAINT_TABLE entry
  baselines/                    # committed oracle outputs (small JSONs)
    word_trpr_pairwise.json     # produced by Phase 2
    word_rpr_deviations.json    # produced by Phase 3
```

### Roundtrip Engine API

```python
@dataclass
class RoundtripResult:
    input_path: Path
    output_path: Path
    repair_dialog_seen: bool
    repair_dialog_text: str | None
    elapsed_seconds: float

def roundtrip(input_docx: Path, *, output_dir: Path | None = None,
              timeout: float = 60.0) -> RoundtripResult: ...
```

Default behavior: stage input in a temp dir, open in Word, dismiss repair dialog as Yes if present, Save through Word, return the saved path along with structured metadata.

### Open + Save Flow

1. **Stage input.** Copy input DOCX to a fresh temp directory. Avoids polluting the source path; Word's auto-save and permission semantics work cleanly on a private file.
2. **Launch Word.** `open -W -a "Microsoft Word"` waits until Word is reachable. Reuse if already running.
3. **Open file via AppleScript Word command.** `tell application "Microsoft Word" to open file POSIX path of "{path}"` — this is more robust than UI scripting; Word's AppleScript dictionary supports it directly.
4. **Wait for document load.** Poll `documents` collection for the file's `name` to appear. Timeout after `timeout` seconds.
5. **Detect repair dialog.** A modal "Microsoft Word" alert dialog with text containing "unreadable content" or "recover the contents". Use System Events UI scripting to find it. If present, click "Yes" and record the dialog text.
6. **Save through Word.** AppleScript `save` command with explicit file format `format word document format`. Saves to a target path inside the temp dir.
7. **Close the document.** AppleScript `close active document saving no` (the explicit Save above already wrote the file; this avoids a second prompt).
8. **Return.** A `RoundtripResult` with both paths, the repair flag, and timing.

### Repair Dialog Handling

- **Detection mechanism.** Poll for an alert window owned by Microsoft Word with title text matching `"unreadable content"`, `"recover the contents"`, `"document.docx contains unreadable"` (configurable list of patterns).
- **Default action: Yes.** Dismissing as Yes preserves the post-repair XML, which is the oracle.
- **Configurable: No.** Optional `accept_repair: bool = True` arg to `roundtrip()` — when False, dismiss as No and the operation fails (the input was not openable). Useful for binary "would Word reject this outright?" testing.
- **Logging.** Capture and return the dialog's full text. Different repair dialogs (e.g., "missing styles" vs "malformed XML") indicate different failure modes.

### Spec 010 Scenario Driver

```python
def run_property_ordering_oracle(
    constraint: ChildSequence,
    output_path: Path,
) -> None:
    """Generate scenarios for one CONSTRAINT_TABLE entry, roundtrip each
    through Word, write a JSON oracle report.

    Scenario set per constraint:
      - baseline: canonical-ordered children only (control case)
      - pairwise swaps: every (i, j) where i < j and i, j are different
        children. Tests every adjacent and non-adjacent reordering.
      - skip-then-restore: `[children[0], children[2], children[1]]` patterns
        — checks whether Word treats subsequence violations the same as
        adjacent swaps.
    """
```

Output JSON shape:

```json
{
  "constraint": "CT_TrPr",
  "spec_section": "ECMA-376 §17.4.79",
  "word_version": "Microsoft Word 16.79.x (build ...)",
  "run_at": "2026-04-28T12:00:00Z",
  "scenarios": [
    {
      "id": "trPr-baseline",
      "input_children_order": ["cnfStyle", "cantSplit", "tblHeader"],
      "post_word_children_order": ["cnfStyle", "cantSplit", "tblHeader"],
      "repair_dialog_seen": false,
      "verdict": "preserved"
    },
    {
      "id": "trPr-swap-cantSplit-tblHeader",
      "input_children_order": ["cnfStyle", "tblHeader", "cantSplit"],
      "post_word_children_order": ["cnfStyle", "cantSplit", "tblHeader"],
      "repair_dialog_seen": true,
      "repair_dialog_text": "The document name.docx contains unreadable content...",
      "verdict": "repaired",
      "diff": "cantSplit moved before tblHeader"
    }
  ]
}
```

### Scenario DOCX Synthesis

For each scenario, the driver builds a minimal valid DOCX containing exactly one instance of the parent property element with the requested children. Reuses the `_build_docx`/`_trpr` test helpers from `tests/test_word_compat_ordering.py` (extracted to a shared location). The synthetic file is valid except for the deliberate property-element ordering; that's the only signal Word should react to.

### XML Diff

After roundtrip, unzip the post-Word DOCX, locate the same property element by parent local name and a stable XPath (or by index — minimal scenarios have only one instance), serialize children's local names, and compare to input. The diff is intentionally child-order-only; attribute changes and added/removed children are reported but treated as separate signals.

## Test Layout

- `tests/test_word_oracle_engine.py` — unit tests for the parts that don't need Word: scenario generation, XML diff helper, report serialization. CI-runnable.
- `tests/test_word_oracle_smoke.py` — marked `requires_word_app`, runs a single end-to-end roundtrip on a known-good DOCX and asserts `preserved` verdict. Skipped in CI; run manually as a sanity check.
- `tools/oracle/baselines/*.json` — committed oracle output from past runs. Treated as data, not code; not regenerated automatically. A separate test asserts these JSONs parse against the schema.

## Acceptance Criteria

1. `tools/oracle/word_roundtrip.py:roundtrip()` opens an arbitrary DOCX in Word, saves through Word, and returns a `RoundtripResult` with the post-Word path. Working on macOS with Microsoft 365 Word installed.
2. The repair dialog detection correctly distinguishes:
   - Clean roundtrip (no dialog, post-Word file is a clean rewrite).
   - Repair dialog seen and accepted (dialog text captured, post-Word file is Word's repaired version).
3. The spec 010 scenario driver produces a complete JSON oracle for `CT_TrPr` pairwise swaps. The baseline (canonical order) reports `preserved`; at least one synthetic deviation reports `repaired`.
4. The oracle's verdict on Shaun's exact issue #3 repro (`tblHeader` before `cantSplit`) is `repaired`.
5. Developer-machine setup is documented in `tools/oracle/README.md`, including the first-run permission prompts and how to recover from a stuck Word session.
6. `pytest` runs cleanly on a machine without Word — Word-dependent tests are skipped via the `requires_word_app` marker.
7. The oracle output is committed to the repo under `tools/oracle/baselines/` and referenced by spec 010 Phase 3/4.

## Rollout Plan

### Phase 1 — Engine (`word_roundtrip.py` + `word_window.py`)

Goal: minimum viable Open + Save through Word with structured repair-dialog reporting. ~250–400 LOC.

Phase exit gate: the engine roundtrips three test DOCX inputs end-to-end (one clean, one with a known repair-triggering ordering, one borderline) and returns correct `RoundtripResult` data. Unit-tested logic separated from Word-dependent flow.

### Phase 2 — Spec 010 driver (`word_repair_oracle.py`) + `CT_TrPr` baseline

Goal: scenario generator for `CT_TrPr` pairwise swaps and skip-then-restore patterns, full oracle run committed under `tools/oracle/baselines/word_trpr_pairwise.json`.

Phase exit gate: oracle JSON is committed, documents `repaired` for at least the issue #3 repro and `preserved` for the canonical baseline. Spec 010 Phase 1 (`CT_TrPr` constraint) is reviewed in light of the oracle data; specific over-flags are removed if found.

### Phase 3 — `rPr` and `pPr` deviation oracle

Goal: feed the corpus-mined deviation patterns from `docs/word_compat/` validation through the roundtripper. Each pattern → `repaired` or `preserved` verdict. Output committed under `tools/oracle/baselines/word_rpr_deviations.json` (and similarly for pPr).

Phase exit gate: oracle data is committed; spec 010 Phase 3 ships `CT_PPr`/`CT_RPr` constraints derived from `preserved` patterns only.

### Phase 4 (optional) — Visual capture

Goal: only if Tier 1 proves insufficient (e.g., Word silently rewrites a file in a way the XML diff misses but a screenshot reveals). Wire in TokenMoulds-style screenshot capture as an additional `RoundtripResult` field.

## Risk Register

1. **Risk:** Word-for-Mac UI changes between versions break dialog detection or save command.
   - Mitigation: abstract the dialog patterns and command surface behind a small adapter; pin a specific Word build in the README and note any drift in `tools/oracle/baselines/` filenames.
2. **Risk:** Word's AppleScript `save` doesn't trigger the same repair flow as opening via UI does.
   - Mitigation: validate during Phase 1 by running the issue #3 repro through both code paths. If `save` bypasses repair, fall back to UI-driven Save (Cmd+S keystroke).
3. **Risk:** Repair dialog has many shapes; we mis-detect or fail to dismiss.
   - Mitigation: keep dialog patterns configurable; log full dialog text on every roundtrip; treat any unhandled dialog as a roundtrip failure rather than silent success.
4. **Risk:** Oracle output is non-reproducible across Word versions.
   - Mitigation: every committed oracle JSON includes the Word version string at run time. Treat the data as a snapshot, not a contract.
5. **Risk:** Roundtripper hangs on a stuck dialog or modal sheet, blocking the developer's machine.
   - Mitigation: hard timeout per roundtrip with cleanup (force-close active document); document recovery in README.
6. **Risk:** Permissions prompts derail first-time setup.
   - Mitigation: README includes a checklist with the exact System Settings panes; provide a small `tools/oracle/preflight.py` that surfaces missing permissions before the engine attempts anything.
7. **Risk:** The oracle scenario set is too narrow (only pairwise swaps), missing failure modes that involve more than two children moving.
   - Mitigation: phase-3 expands the scenario generators with skip-then-restore and partial-reverse patterns; document the coverage in each baseline JSON.

## Open Questions

- Does Word for Mac's AppleScript `save` command trigger the same repair flow as a UI-initiated Save? Requires Phase 1 verification.
- Should the engine record the post-Word DOCX in addition to the diff? Useful for postmortem but adds bytes to the repo; recommend committing only the JSON and keeping post-Word files in `.gitignored` scratch storage.
- For multi-instance scenarios (e.g., a paragraph with two `<w:rPr>` elements, each with a different ordering), how does the diff attribute results? Recommend: scenarios are scoped to a single property-element instance per file in v1 to avoid this.
- Word may treat the document differently depending on `<?mso-application?>` processing instructions or attached templates. Recommend: synthesize scenarios with no template attachments and a fixed minimal `[Content_Types].xml`; document the assumption.

## Definition of Done

All must be true:

1. The Tier 1 roundtripper engine exists at `tools/oracle/word_roundtrip.py` and successfully roundtrips an arbitrary DOCX through Word for Mac on a configured macOS machine.
2. Repair dialogs are detected, accepted by default, and reported in the structured result.
3. The spec 010 scenario driver produces a complete JSON oracle for `CT_TrPr` pairwise swaps; the JSON is committed under `tools/oracle/baselines/`.
4. Spec 010 Phase 1's `CT_TrPr` constraint is reviewed against the oracle, and any over-flags are corrected.
5. The same oracle is run for the `rPr`/`pPr` corpus deviation patterns; data committed; spec 010 Phase 3 lands with empirically-validated constraints.
6. Developer-machine setup is documented; `pytest` is green on machines without Word; Word-dependent tests are properly marked.
7. The `tools/oracle/` directory is structured so the next Word-app-compat oracle (font fallback, MC alternate content, etc.) can plug in as a sibling driver without modifying the engine.
