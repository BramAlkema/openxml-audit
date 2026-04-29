"""AppleScript / osascript helpers for Microsoft Word for Mac.

Pattern reference: TokenMoulds' `tools/visual/pptx_window.py`. The Word
surface is similar in shape but the AppleScript dictionary differs
(Word exposes `documents`, `active document`, and a `save as` command
with explicit format constants; PowerPoint uses a different vocabulary).

Spec: `specs/011-word-roundtrip-oracle.md` Phase 1.
"""

from __future__ import annotations

import json
import shlex
import subprocess
from pathlib import Path

_WORD_PROCESS_NAME = "Microsoft Word"
_WORD_APP_BUNDLE = Path("/Applications/Microsoft Word.app")

# Repair-dialog text patterns. These are the substrings we look for in
# the body of a modal alert to decide that Word triggered its
# "unreadable content" recovery flow. The list is intentionally
# generous — different Word versions and different repair scenarios
# produce different exact strings.
REPAIR_DIALOG_PATTERNS = (
    "unreadable content",
    "recover the contents",
    "contains unreadable",
    "errors were detected",
    "do you want to recover",
    # The "harder failure" dialog: Word can't open the file at all rather
    # than offering to repair. Different from the unreadable-content dialog
    # but for our oracle purposes both signal "Word rejected the input."
    "experienced an error",
    "error trying to open",
    # The "this file is corrupt — open and repair?" dialog. First seen
    # on 2026-04-29 baseline collection runs against TokenMoulds-emitted
    # scratch .docx files that Word refuses to open in their as-emitted
    # form. Wording observed: "Word was unable to read this document. It
    # may be corrupt. Try one or more of the following: Open and Repair
    # the file. Open the file with the Text Recovery converter."
    "unable to read this document",
    "may be corrupt",
    "open and repair",
    "text recovery converter",
)


class OsascriptError(RuntimeError):
    """Raised when an osascript invocation returns non-zero or times out."""


def _word_app_target() -> str:
    """Return the value to pass to `open -a` — bundle path if available,
    process name otherwise."""
    return str(_WORD_APP_BUNDLE) if _WORD_APP_BUNDLE.exists() else _WORD_PROCESS_NAME


def _applescript_quote(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def osascript(script: str, *, timeout: float | None = 30.0) -> str:
    """Run an AppleScript via osascript, return trimmed stdout, raise on
    non-zero. Mirror of TokenMoulds' helper for consistency."""
    result = subprocess.run(
        ["osascript", "-"],
        input=script,
        text=True,
        capture_output=True,
        check=False,
        timeout=timeout,
    )
    if result.returncode != 0:
        raise OsascriptError(result.stderr.strip() or "osascript failed")
    return result.stdout.strip()


def osascript_jxa(script: str, *, timeout: float | None = 30.0) -> str:
    """Run a JXA (JavaScript) script via osascript -l JavaScript."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-"],
        input=script,
        text=True,
        capture_output=True,
        check=False,
        timeout=timeout,
    )
    if result.returncode != 0:
        raise OsascriptError(result.stderr.strip() or "osascript failed")
    return result.stdout.strip()


def launch_word_app() -> None:
    """Launch Microsoft Word (no-op if already running)."""
    target = _word_app_target()
    launch_cmd = f"open -W -a {shlex.quote(target)}"
    subprocess.Popen(  # noqa: S603 — controlled command, no shell metacharacters
        ["/bin/zsh", "-lc", launch_cmd],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )


def is_word_running() -> bool:
    """Return True if the Word process is currently active."""
    try:
        process_name = _applescript_quote(_WORD_PROCESS_NAME)
        result = osascript(
            'tell application "System Events" to '
            f"return (count of (processes whose name is {process_name})) > 0",
            timeout=5.0,
        )
    except (OsascriptError, subprocess.TimeoutExpired):
        return False
    return result.lower() == "true"


def word_version() -> str | None:
    """Return Word's reported version string, or None if Word is unreachable."""
    try:
        return osascript(
            f'tell application {_applescript_quote(_WORD_PROCESS_NAME)} to return version',
            timeout=10.0,
        ) or None
    except (OsascriptError, subprocess.TimeoutExpired):
        return None


def open_document(path: Path, *, timeout: float = 15.0) -> bool:
    """Dispatch a DOCX open to Word via AppleScript. Returns True if the
    `osascript` call returned cleanly, False if it timed out or errored —
    callers should poll `is_document_open()` regardless, since the open
    may complete asynchronously even when the AppleScript dispatch hangs.

    NB: omits `activate` because the activate command can block on Word's
    welcome / Recent Files window. We rely on `launch_word_app` having
    already brought Word foreground.
    """
    posix = str(path.resolve())
    script = f"""
tell application {_applescript_quote(_WORD_PROCESS_NAME)}
    open (POSIX file {_applescript_quote(posix)})
end tell
"""
    try:
        osascript(script, timeout=timeout)
    except (OsascriptError, subprocess.TimeoutExpired):
        return False
    return True


def list_open_document_names() -> list[str]:
    """Return the display names of every open document."""
    try:
        out = osascript(
            f'tell application {_applescript_quote(_WORD_PROCESS_NAME)} to '
            'return name of every document as string',
            timeout=10.0,
        )
    except (OsascriptError, subprocess.TimeoutExpired):
        return []
    if not out:
        return []
    # AppleScript joins lists with ", " by default when coerced to string.
    return [name.strip() for name in out.split(",") if name.strip()]


def _matching_document_script(path: Path) -> str:
    """AppleScript fragment that exposes a `documentMatches` handler and
    sets `targetPosix`/`targetHfs`/`targetName` for the given path.

    Mirrors TokenMoulds' robust multi-representation matching: Word may
    report a document's identity via `full name`, `path & name`, or just
    `name`, depending on how it was opened.

    NB: callers must use index-based iteration (`repeat with i from 1 to
    count`), not `repeat with d in documents` — the latter causes
    infinite loops inside nested tells.
    """
    target_posix = json.dumps(str(path.resolve()))
    target_name = json.dumps(path.name)
    return f"""
set targetPosix to {target_posix}
set targetName to {target_name}
set targetHfs to ""
try
    set targetHfs to ((POSIX file targetPosix) as text)
end try

on documentMatches(docRef, targetPosix, targetHfs, targetName)
    tell application {_applescript_quote(_WORD_PROCESS_NAME)}
    try
        set docFullName to (full name of docRef) as text
        if docFullName is targetPosix then
            return true
        end if
        if targetHfs is not "" and docFullName is targetHfs then
            return true
        end if
    end try
    try
        set docPath to (path of docRef) as text
        set docName to (name of docRef) as text
        set combinedPath to docPath & docName
        if combinedPath is targetPosix then
            return true
        end if
        if targetHfs is not "" and combinedPath is targetHfs then
            return true
        end if
    end try
    try
        if ((name of docRef) as text) is targetName then
            return true
        end if
    end try
    end tell
    return false
end documentMatches
"""


def is_document_open(path: Path) -> bool:
    """Return True if Word currently has the given file open as a document.

    Tries multiple identity representations (POSIX path, HFS path, name)
    because Word reports identity differently depending on whether the
    file was opened directly, via Recent Files, or by drag-and-drop.
    """
    script = (
        _matching_document_script(path)
        + f"""
tell application {_applescript_quote(_WORD_PROCESS_NAME)}
    try
        set docCount to count of documents
        repeat with i from 1 to docCount
            set docRef to document i
            if my documentMatches(docRef, targetPosix, targetHfs, targetName) then
                return "true"
            end if
        end repeat
    end try
end tell
return "false"
"""
    )
    try:
        return osascript(script, timeout=10.0).strip().lower() == "true"
    except (OsascriptError, subprocess.TimeoutExpired):
        return False


def activate_document(path: Path) -> bool:
    """If `path` is open in Word, set it as the active document. Returns
    True on success."""
    script = (
        _matching_document_script(path)
        + f"""
tell application {_applescript_quote(_WORD_PROCESS_NAME)}
    try
        set docCount to count of documents
        repeat with i from 1 to docCount
            set docRef to document i
            if my documentMatches(docRef, targetPosix, targetHfs, targetName) then
                activate document i
                return "true"
            end if
        end repeat
    end try
end tell
return "false"
"""
    )
    try:
        return osascript(script, timeout=10.0).strip().lower() == "true"
    except (OsascriptError, subprocess.TimeoutExpired):
        return False


def close_active_document_saving() -> None:
    """Close the active document, saving to its current path.

    Word's M365 sdef advertises Cocoa-standard `save` support on the
    window class but empirically returns -1708. The same is true of
    Word's bespoke `save as`. The `close ... saving yes` variant goes
    through `handleCloseScriptCommand:` which IS wired up in practice.
    Word writes the modified document to its current path before closing.

    Caller responsibility: ensure the document was opened from a path
    the caller controls (i.e., a copy in our staging directory). After
    this call, that path holds Word's post-roundtrip XML.
    """
    script = f"""
tell application {_applescript_quote(_WORD_PROCESS_NAME)}
    if (count of documents) > 0 then
        close active document saving yes
    end if
end tell
"""
    osascript(script, timeout=30.0)


def close_active_document() -> None:
    """Close the active document without saving (the explicit save above
    already wrote the file)."""
    script = f"""
tell application {_applescript_quote(_WORD_PROCESS_NAME)}
    if (count of documents) > 0 then
        close active document saving no
    end if
end tell
"""
    osascript(script, timeout=15.0)


def find_repair_dialog_text(patterns: tuple[str, ...] = REPAIR_DIALOG_PATTERNS) -> str | None:
    """If a modal alert is currently presented by Word and its body text
    matches any of `patterns` (case-insensitive substring), return the
    full text. Otherwise return None.

    Uses System Events UI scripting; relies on Word presenting the alert
    as a child sheet/window of its main frame. Tested empirically; subject
    to Word UI changes.
    """
    script = f"""
tell application "System Events"
    if not (exists process {_applescript_quote(_WORD_PROCESS_NAME)}) then return ""
    tell process {_applescript_quote(_WORD_PROCESS_NAME)}
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
        text = osascript(script, timeout=10.0)
    except (OsascriptError, subprocess.TimeoutExpired):
        return None
    if not text:
        return None
    lowered = text.lower()
    for pattern in patterns:
        if pattern.lower() in lowered:
            return text.strip()
    return None


def click_dialog_button(button_label: str) -> bool:
    """Click a button by label in any sheet/dialog currently presented by
    Word. Returns True if the button was found and clicked."""
    script = f"""
tell application "System Events"
    if not (exists process {_applescript_quote(_WORD_PROCESS_NAME)}) then return "false"
    tell process {_applescript_quote(_WORD_PROCESS_NAME)}
        try
            repeat with w in windows
                try
                    set sh to sheets of w
                    repeat with s in sh
                        try
                            click button {_applescript_quote(button_label)} of s
                            return "true"
                        end try
                    end repeat
                end try
                try
                    click button {_applescript_quote(button_label)} of w
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
    except (OsascriptError, subprocess.TimeoutExpired):
        return False
    return result.strip().lower() == "true"
