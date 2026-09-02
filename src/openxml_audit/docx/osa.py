"""Word-specific AppleScript / JXA helpers for UI-driven opens and saves.

Builds on `openxml_audit.osa` primitives. macOS + Microsoft Word only.

This module is the canonical home for Word window primitives —
parallel to `openxml_audit.pptx.osa` (since 0.6.4) and
`openxml_audit.xlsx.osa` (since 0.6.5). Spec 030 (0.7.4) consolidated
the Word primitives that historically lived in
`tools/oracle/word_window.py` (since Spec 011 / 0.5.0) into this
module so the four-format oracle ladder uses one consistent in-package
osa layer. `tools/oracle/word_window.py` now re-exports from here for
back-compat with consumers that imported it directly.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

from openxml_audit.osa import launch_app, osascript

__all__ = [
    "REPAIR_DIALOG_ACCEPT_BUTTON_LABELS",
    "REPAIR_DIALOG_PATTERNS",
    "WORD_APP_BUNDLE",
    "WORD_APP_ID",
    "WORD_PROCESS_NAME",
    "activate_document",
    "click_dialog_button",
    "close_active_document",
    "close_active_document_saving",
    "close_document",
    "close_document_saving",
    "dismiss_any_leftover_modal",
    "dismiss_repair_dialog",
    "find_repair_dialog_text",
    "is_document_open",
    "is_word_running",
    "launch_word",
    "launch_word_app",
    "list_open_document_names",
    "open_document",
    "save_document",
    "word_version",
]


WORD_APP_ID = "com.microsoft.Word"
WORD_PROCESS_NAME = "Microsoft Word"
WORD_APP_BUNDLE = Path("/Applications/Microsoft Word.app")


# Repair-dialog text patterns. Word's "unreadable content" repair
# flow has multiple variants (soft repair offering recovery, hard
# error refusing to open). The list is intentionally generous —
# different Word builds and different repair categories produce
# different exact strings. Cross-checked against real dialogs
# observed during 0.6.6 / 0.6.9 baseline runs.
REPAIR_DIALOG_PATTERNS: tuple[str, ...] = (
    "unreadable content",
    "recover the contents",
    "contains unreadable",
    "errors were detected",
    "do you want to recover",
    # The "harder failure" dialog: Word can't open the file at all.
    "experienced an error",
    "error trying to open",
    # The "this file is corrupt — open and repair?" dialog. First
    # observed during 0.6.6 baseline collection on TokenMoulds-emitted
    # scratch .docx files.
    "unable to read this document",
    "may be corrupt",
    "open and repair",
    "text recovery converter",
)


# Button labels to try (in order) when accepting Word's repair
# dialog. "Yes" is the legacy primary affirmative; "OK" / "Recover"
# / "Open and Repair" cover the variants observed on Office for Mac
# M365 since Spec 011.
REPAIR_DIALOG_ACCEPT_BUTTON_LABELS: tuple[str, ...] = (
    "Yes",
    "OK",
    "Recover",
    "Open and Repair",
    "Open",
)


def _word_app_target() -> str:
    """Return the value to pass to `open -a` — bundle path if available,
    process name otherwise."""
    return str(WORD_APP_BUNDLE) if WORD_APP_BUNDLE.exists() else WORD_PROCESS_NAME


def _applescript_quote(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def launch_word() -> None:
    """Launch Word.app (via `open -W`). Idempotent."""
    launch_app(WORD_APP_BUNDLE, WORD_PROCESS_NAME)


# Spec 011 / `tools/oracle/word_window.py` historical name. Same as
# `launch_word`; kept as alias for back-compat.
launch_word_app = launch_word


def is_word_running() -> bool:
    """True if a Word process is currently visible to System Events."""
    script = f"""
tell application "System Events"
    return (exists process "{WORD_PROCESS_NAME}")
end tell
"""
    try:
        out = osascript(script, timeout=5.0)
    except (RuntimeError, subprocess.TimeoutExpired):
        return False
    return out.lower() == "true"


def word_version() -> str | None:
    """Return Word's `version` property, or None if Word isn't reachable.

    Used in `RoundtripResult.word_version` so observation reports
    record what build the run executed against.
    """
    try:
        out = osascript(
            f'tell application "{WORD_PROCESS_NAME}" to return version',
            timeout=5.0,
        )
    except (RuntimeError, subprocess.TimeoutExpired):
        return None
    return out or None


def open_document(
    docx_path: Path,
    *,
    timeout: float | None = 30.0,
) -> bool:
    """Tell Word to open `docx_path`. Returns True on success.

    Tries `open -b com.microsoft.Word` first, then `open -a "Microsoft
    Word"` as a fallback. Both use macOS's `open` command which
    handles launching Word if it isn't running.

    Returns False (rather than raising) if all `open` invocations
    fail — corpus walks need to record "failed to open" as data,
    not as a stalled run.
    """
    docx_path = docx_path.resolve()
    launch_timeout = max(5.0, min(timeout or 30.0, 15.0))
    for command in (
        ["open", "-b", WORD_APP_ID, str(docx_path)],
        ["open", "-a", WORD_PROCESS_NAME, str(docx_path)],
    ):
        try:
            subprocess.run(
                command,
                check=True,
                timeout=launch_timeout,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
        except Exception:
            continue
    return False


def list_open_document_names() -> list[str]:
    """Return the display names of every open Word document.

    Uses `name of every document as string` rather than
    `repeat with d in documents / name of d` — the latter causes
    infinite loops inside nested `tell` blocks on Office for Mac
    M365 (a Word AppleScript bridge bug, see Spec 022 baseline
    notes). Catches `subprocess.TimeoutExpired` so a not-running
    Word doesn't hang the test suite past the 10s budget (same fix
    we landed for PPTX/XLSX in 0.6.5).
    """
    try:
        out = osascript(
            f'tell application "{WORD_PROCESS_NAME}" to '
            "return name of every document as string",
            timeout=10.0,
        )
    except (RuntimeError, subprocess.TimeoutExpired):
        return []
    if not out:
        return []
    return [name.strip() for name in out.split(",") if name.strip()]


def is_document_open(docx_path: Path) -> bool:
    """True if a document matching `docx_path.name` is currently open in Word.

    Word reports document identity via `full name`, `path & name`, or
    just `name` depending on how it was opened. This implementation
    checks against `name` only — sufficient for the oracle's
    staged-input pattern where the caller knows the staged filename
    is unique among open docs.
    """
    target_name = docx_path.name
    return any(
        name == target_name or name.endswith("/" + target_name) or name.endswith("\\" + target_name)
        for name in list_open_document_names()
    )


def activate_document(docx_path: Path) -> bool:
    """Make the document matching `docx_path.name` Word's active document.

    Word's `close active document saving yes` writes whatever
    document is active; if the staged document isn't activated first,
    a different document might get saved instead. Returns True on
    success.
    """
    name = docx_path.name
    script = f"""
tell application "{WORD_PROCESS_NAME}"
    set targetName to {_applescript_quote(name)}
    repeat with i from 1 to count documents
        set d to document i
        if (name of d) is targetName then
            activate
            set active document to d
            return "true"
        end if
    end repeat
    return "false"
end tell
"""
    try:
        out = osascript(script, timeout=10.0)
    except (RuntimeError, subprocess.TimeoutExpired):
        return False
    return out.strip().lower() == "true"


def save_document(*, timeout: float | None = 30.0) -> None:
    """Save the active Word document via Cmd-S (overwrite in place)."""
    script = f"""
tell application "System Events"
    tell process "{WORD_PROCESS_NAME}"
        set frontmost to true
        delay 0.2
        keystroke "s" using {{command down}}
        delay 0.5
    end tell
end tell
"""
    osascript(script, timeout=timeout)


def close_document(*, timeout: float | None = 30.0) -> None:
    """Close the active Word document without saving."""
    script = f"""
tell application "{WORD_PROCESS_NAME}"
    if (count documents) > 0 then
        close active document saving no
    end if
end tell
"""
    osascript(script, timeout=timeout)


# Spec 011 historical name; same behavior as close_document.
close_active_document = close_document


def close_document_saving(*, timeout: float | None = 30.0) -> None:
    """Close the active document, persisting via the close handler.

    Word's `M365` sdef advertises Cocoa-standard `save` on the
    window class but empirically returns -1708. The same is true of
    Word's bespoke `save as`. The `close ... saving yes` variant
    goes through `handleCloseScriptCommand:` which IS wired up in
    practice — Word writes the modified document to its current
    path before tearing down the window.

    Caller responsibility: ensure the document was opened from a
    path the caller controls (i.e. a copy in our staging directory).
    After this call, that path holds Word's post-roundtrip XML.
    """
    script = f"""
tell application "{WORD_PROCESS_NAME}"
    if (count documents) > 0 then
        close active document saving yes
    end if
end tell
"""
    osascript(script, timeout=timeout)


# Spec 011 historical name; same behavior as close_document_saving.
close_active_document_saving = close_document_saving


def find_repair_dialog_text(
    patterns: tuple[str, ...] = REPAIR_DIALOG_PATTERNS,
) -> str | None:
    """If a modal alert is currently presented by Word and its body
    text matches any of `patterns` (case-insensitive substring),
    return the full text. Otherwise None.

    Uses System Events UI scripting; relies on Word presenting the
    alert as a child sheet/window of its main frame. Tested
    empirically; subject to Word UI changes.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{WORD_PROCESS_NAME}") then return ""
    tell process "{WORD_PROCESS_NAME}"
        set out to ""
        try
            repeat with w in windows
                try
                    set sh to sheets of w
                    repeat with s in sh
                        try
                            set sts to static texts of s
                            repeat with t in sts
                                set out to out & " " & (value of t as string)
                            end repeat
                        end try
                    end repeat
                end try
                try
                    set sts to static texts of w
                    repeat with t in sts
                        set out to out & " " & (value of t as string)
                    end repeat
                end try
            end repeat
        end try
        return out
    end tell
end tell
"""
    try:
        text = osascript(script, timeout=5.0)
    except (RuntimeError, subprocess.TimeoutExpired):
        return None
    if not text:
        return None
    haystack = text.lower()
    for needle in patterns:
        if needle.lower() in haystack:
            return text
    return None


def click_dialog_button(button_label: str) -> bool:
    """Click a button by exact label in any sheet/dialog currently
    presented by Word. Returns True if the button was found and
    clicked. Mirrors `pptx.osa.click_dialog_button` /
    `xlsx.osa.click_dialog_button`.

    Uses System Events UI scripting; the `repeat with X in
    COLLECTION` idiom is fine here — the hang documented in
    `list_open_document_names` is specific to `tell application
    "Microsoft Word"`, not System Events.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{WORD_PROCESS_NAME}") then return "false"
    tell process "{WORD_PROCESS_NAME}"
        try
            repeat with w in windows
                try
                    set sh to sheets of w
                    repeat with s in sh
                        try
                            click button "{button_label}" of s
                            return "true"
                        end try
                    end repeat
                end try
                try
                    click button "{button_label}" of w
                    return "true"
                end try
            end repeat
        end try
        return "false"
    end tell
end tell
"""
    try:
        result = osascript(script, timeout=10.0)
    except (RuntimeError, subprocess.TimeoutExpired):
        return False
    return result.strip().lower() == "true"


def dismiss_repair_dialog(
    *,
    accept_labels: tuple[str, ...] = REPAIR_DIALOG_ACCEPT_BUTTON_LABELS,
) -> tuple[bool, str | None, str | None]:
    """Detect and click through Word's repair dialog.

    Returns `(was_seen, dialog_text, clicked_label)`. If no dialog is
    visible, returns `(False, None, None)`. If a dialog is visible
    but no `accept_labels` button could be clicked, returns
    `(True, dialog_text, None)` — the caller decides whether to
    treat that as a hard failure.
    """
    text = find_repair_dialog_text()
    if text is None:
        return False, None, None
    for label in accept_labels:
        if click_dialog_button(label):
            return True, text, label
    return True, text, None


def dismiss_any_leftover_modal(*, timeout: float | None = 5.0) -> bool:
    """Best-effort dismissal of any leftover modal sheet via Escape.

    Word may show a follow-up info modal after the primary repair
    dialog or after close-with-save returns. Sending Escape to the
    frontmost Word window discards the modal in one shot. Symmetric
    with `pptx.osa.dismiss_any_leftover_modal` and
    `xlsx.osa.dismiss_any_leftover_modal`.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{WORD_PROCESS_NAME}") then return "false"
    tell process "{WORD_PROCESS_NAME}"
        try
            key code 53
            return "true"
        end try
        return "false"
    end tell
end tell
"""
    try:
        out = osascript(script, timeout=timeout)
    except (RuntimeError, subprocess.TimeoutExpired):
        return False
    return out.strip().lower() == "true"
