# macOS Setup for Roundtrip Oracles

The `tools/oracle/*_repair_oracle.py` family drives Microsoft Word, Excel, PowerPoint, and LibreOffice through automation. macOS gates that automation behind two privacy categories. This page tells you exactly what to grant.

If you skip this and try to run the oracles cold, the AppleScript calls will hang for ~10 seconds and return empty results, the smoke tests will skip with timeout messages, and you'll wonder why nothing's happening. Granting the permissions is a one-time setup.

## Quick check

```bash
python -m tools.oracle.preflight
```

Expected output when everything is ready:

```
  word       OK
  excel      OK
  powerpoint OK
  odf        OK
```

If any line says `FAIL`, the issue list points at what to fix.

## What to grant

### 1. Automation (control Office apps via AppleScript)

`System Settings` → `Privacy & Security` → `Automation`.

Find your terminal app in the list (whatever launches `python` — typically `iTerm` or `Terminal`). Expand it. Tick:

- Microsoft Word
- Microsoft Excel
- Microsoft PowerPoint

If your terminal app isn't in the list yet, run `python -m tools.oracle.preflight --engine word` once. The first AppleScript call triggers macOS to register your terminal in the Automation panel; then the panel will show it.

### 2. Accessibility (System Events keystrokes)

Same panel: `Privacy & Security` → `Accessibility`.

Tick the same terminal app. This is required because the `save_workbook` / `save_document` / `save_presentation` paths route through System Events `keystroke "s" using {command down}`, and System Events keystrokes are an Accessibility-gated capability.

### 3. (No setup needed) LibreOffice

The ODF oracle uses headless `soffice --convert-to`, not AppleScript. No Privacy & Security grants involved. Just have LibreOffice installed at `/Applications/LibreOffice.app` (or `soffice` on PATH).

## When permissions are denied or revoked

Symptoms:

- `osascript` returns `execution error: Not authorized to send Apple events to ...` (`error -1743`)
- The preflight reports `Could not query <App>'s version via AppleScript`
- `list_open_workbook_names()` / `list_open_presentation_names()` returns `[]` even though apps are running
- Smoke tests (`test_*_roundtrip_oracle.py::test_list_open_*_returns_list`) skip with timeout messages

Fix: re-tick the app under Automation, then re-run preflight.

## Sandbox staging directories

Office apps run under macOS's App Sandbox. They have read/write access to `~/Documents` by default; tmpfs paths (`/tmp`, `/var/folders/...`) require explicit `Full Disk Access` for each app. The oracles default to `~/Documents/.{word,xlsx,pptx}_oracle_runs/<id>/` to avoid the issue.

Override (e.g. for CI on a dedicated machine with Full Disk Access):

```bash
WORD_ORACLE_STAGE=/tmp/word_runs
XLSX_ORACLE_STAGE=/tmp/xlsx_runs
PPTX_ORACLE_STAGE=/tmp/pptx_runs
```

The ODF oracle's per-call `UserInstallation` profile dir uses `tempfile.mkdtemp` and is unaffected (LibreOffice's headless mode isn't sandboxed).

## Why this is operationally fragile

The macOS automation surface was designed for occasional UI-driven scripting, not corpus-walking thousands of files. Symptoms you'll hit at scale:

- AppleScript dispatcher races during cold app launch — fixed by `time.sleep(1.0)` after `launch_*` (already in each oracle).
- Office apps presenting modal sheets the AppleScript can't dismiss — Word and PowerPoint repair dialogs (Spec 011 + Spec 020) detect these; auto-dismiss is Phase 2 work.
- App Sandbox blocking access to a path you've previously approved (`Documents` re-permission prompts after macOS updates) — re-grant under `Files and Folders`.
- Privacy permissions silently dropped after Office app updates — re-tick under Automation.

Rule of thumb: when an oracle starts producing `open_failed` outcomes en masse, run preflight before debugging the oracle code. The cause is almost always permission state, not logic.
