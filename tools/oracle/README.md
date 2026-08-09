# Oracle Tools

This directory contains the app and service-backed oracle engines used by
`openxml-audit-oracle`. It ships in the wheel because the packaged dispatcher
imports these modules; individual engines still have their own platform,
application, network, and credential requirements.

The Euro-Office conversion engine is service-backed and cross-platform:

```bash
openxml-audit-oracle eurooffice https://files.example.test/sample.odt
```

It reads `EUROOFFICE_ORACLE_URL` and (when enabled by the server)
`EUROOFFICE_ORACLE_JWT_SECRET`. See
[`specs/038-eurooffice-conversion-oracle.md`](../../specs/038-eurooffice-conversion-oracle.md).

## Word Roundtrip Oracle

Spec: [`specs/011-word-roundtrip-oracle.md`](../../specs/011-word-roundtrip-oracle.md)

This is **developer-machine infrastructure**. It runs Microsoft Word for
Mac via osascript, opens a DOCX, saves it through Word, and returns the
post-Word file. The diff between input and post-Word is the empirical
oracle for any "would Word repair this?" question.

The Word engine is developer-machine infrastructure: although its module ships
for dispatcher availability, it needs Microsoft Word for Mac and cannot run in
ordinary Linux CI.

## Why It Exists

Spec 010 (Word compat element ordering) flags property-element reorderings
that *probably* trigger Word's repair dialog. The certainty rests on
proxies (XSDs, SDK schema metadata, real-world DOCX corpora). Each
proxy has been wrong somewhere. The roundtrip oracle gives us direct
ground truth: open the file in Word, save, diff. Any change to the
serialized property element is something Word repaired.

## Requirements

- macOS (osascript-driven; no Linux or Windows support in v1)
- Microsoft Word for Mac (Microsoft 365 / 16.x recommended)
- System Settings → Privacy & Security:
  - **Automation**: allow your terminal (Terminal, iTerm2, VS Code, …)
    to control "Microsoft Word"
  - **Accessibility**: allow your terminal to control your computer
    (needed for UI scripting against the repair dialog)
  - **Files & Folders** *or* **Full Disk Access** for "Microsoft Word":
    Word for Mac is App Sandboxed and cannot open files outside its
    granted scope. The roundtripper stages inputs in
    `~/Documents/.word_oracle_runs/` by default, which Word's sandbox
    accepts without prompting. If you want to stage elsewhere (set via
    `WORD_ORACLE_STAGE` environment variable), grant Word **Full Disk
    Access**.
  - **Screen Recording**: only required if Phase 4 visual capture is
    enabled (currently deferred)

### Why Documents/

The first symptom of missing file-access permissions is the AppleScript
`open` command hanging silently — Word is waiting for a sandbox-permission
response it never displays. Staging in `~/Documents/...` avoids that
entirely. This is the same pattern Office's own automation hooks use.

## First-Run Setup

Run the preflight check before anything else:

```bash
python -m tools.oracle.preflight
```

It surfaces missing prerequisites with actionable error messages. If the
first AppleScript dispatch is rejected, macOS will prompt for
Automation permission — accept it, then re-run preflight until it
prints `OK`.

## Quick Start

```python
from pathlib import Path
from tools.oracle.word_roundtrip import roundtrip

result = roundtrip(Path("scenario.docx"))
print(result.repair_dialog_seen, result.output_path)
```

`result.output_path` is the post-Word DOCX. Compare its property-element
serialization to the input to derive the oracle verdict.

For batch / scenario-driven oracle runs, see the spec-010 driver
documented in spec 011 Phase 2 (`tools/oracle/word_repair_oracle.py` —
not yet implemented).

## Recovery From a Stuck Word Session

If a roundtrip leaves Word in a broken state (modal dialog Word can't
dismiss, frozen save sheet, etc.):

```bash
osascript -e 'tell application "Microsoft Word" to quit saving no'
# or, if that hangs:
killall "Microsoft Word"
```

Then delete any leftover staging directories:

```bash
rm -rf /tmp/word_oracle_*
```

## Status

Phase 1 of spec 011 is in this commit: engine, preflight, smoke tests,
and one immediately surprising empirical finding (below). The downstream
`tools/oracle/word_repair_oracle.py` (spec 010 scenario driver) is
pending Phase 2.

### What the engine does

`tools/oracle/word_roundtrip.py:roundtrip(docx)`:
1. Stages the input under `~/Documents/.word_oracle_runs/run-XXXXXXXX/`,
   keeping `original_<name>.docx` as a reference and a `<name>.docx`
   working copy
2. Launches Word and dispatches `open (POSIX file ...)` via AppleScript
3. Polls Word's `documents` collection (using TokenMoulds-style
   multi-representation matching) until the document is registered
4. While polling, watches for "unreadable content" repair dialogs and
   dismisses with Yes by default (configurable)
5. Calls `close active document saving yes` — Word's bespoke `save as`
   and the inherited Cocoa `save` are advertised in Word.sdef but both
   return -1708 in practice; close-with-save is the only AppleScript
   persist path that works in Word for Mac M365
6. Returns a `RoundtripResult` with input / staged-working / output
   paths, the repair-dialog flag, and Word's version string

The diff between `original_<name>.docx` and the post-Word working copy
is the oracle verdict.

### First empirical finding: Word M365 may not repair issue #3's repro

The smoke-test run on Word for Mac M365 (16.89.1) shows that
`<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>` — the exact pattern
issue #3 reports as triggering Word's repair dialog — was **preserved
as-is** through the AppleScript open + close-with-save path. No repair
dialog, no XML changes.

That contradicts the assumption underlying spec 010 Phase 1's `CT_TrPr`
constraint. Three plausible explanations to investigate before we trust
or distrust the constraint:

1. **AppleScript open is more permissive than UI File>Open.** The UI
   path may run integrity checks the scripted path skips. Mitigation:
   add a UI-driven open variant to the engine and re-test.
2. **Synthetic minimal DOCX bypasses validation that richer content
   triggers.** Word may only run repair-dialog logic when the document
   has substantial style/theme/relationship structure. Mitigation:
   re-test with a python-docx-generated DOCX that has the trPr planted
   inside a real table.
3. **Word version drift.** Issue #3 was reported on M365; we're on
   16.89.1. The reporter's version may differ. Mitigation: get an exact
   build string from the reporter.

Until at least one of these is investigated, spec 010's `CT_TrPr`
constraint is unverified — possibly a false-positive generator on
Word for Mac M365. The Phase 2 scenario driver should produce the
oracle baseline that resolves this.

This is exactly the kind of surprise the oracle was built to surface.
