"""Excel-specific AppleScript / JXA helpers for UI-driven opens and saves.

Builds on `openxml_audit.osa` primitives. macOS + Microsoft Excel only.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

from openxml_audit.osa import launch_app, osascript

__all__ = [
    "EXCEL_APP_BUNDLE",
    "EXCEL_APP_ID",
    "EXCEL_PROCESS_NAME",
    "REPAIR_DIALOG_ACCEPT_BUTTON_LABELS",
    "REPAIR_DIALOG_PATTERNS",
    "click_dialog_button",
    "close_workbook",
    "close_workbook_saving",
    "dismiss_any_leftover_modal",
    "dismiss_repair_dialog",
    "find_repair_dialog_text",
    "is_workbook_open",
    "launch_excel",
    "list_open_workbook_names",
    "open_workbook",
    "save_workbook",
]


EXCEL_APP_ID = "com.microsoft.Excel"
EXCEL_PROCESS_NAME = "Microsoft Excel"
EXCEL_APP_BUNDLE = Path("/Applications/Microsoft Excel.app")


# Repair-dialog text patterns. Excel surfaces a "We found a problem
# with some content in" / "we can try to recover" modal when it
# repairs a workbook on open. Generous match list — different builds
# and different repair categories produce different exact strings.
REPAIR_DIALOG_PATTERNS: tuple[str, ...] = (
    "found a problem",
    "found unreadable",
    "unreadable content",
    "repair",
    "recover",
    "could not be opened",
    "experienced an error",
    "error trying to open",
    "errors were detected",
)


# Button labels to try (in order) when accepting Excel's repair dialog.
# Excel's "Do you want us to try to recover as much as we can? If you
# trust the source of this workbook, click Yes." dialog shows
# Yes / No. Excel's "found a problem" prompt sometimes uses
# Repair / Cancel. Order is most-specific affirmative first.
REPAIR_DIALOG_ACCEPT_BUTTON_LABELS: tuple[str, ...] = (
    "Yes",
    "Repair",
    "Recover",
    "OK",
    "Open",
)


def launch_excel() -> None:
    """Launch Excel.app (via `open -W`)."""
    launch_app(EXCEL_APP_BUNDLE, EXCEL_PROCESS_NAME)


def open_workbook(
    xlsx_path: Path,
    *,
    timeout: float | None = 30.0,
) -> None:
    """Open a `.xlsx` in Excel via `open -b` / `open -a`."""
    xlsx_path = xlsx_path.resolve()
    launch_timeout = max(5.0, min(timeout or 30.0, 15.0))
    for command in (
        ["open", "-b", EXCEL_APP_ID, str(xlsx_path)],
        ["open", "-a", EXCEL_PROCESS_NAME, str(xlsx_path)],
    ):
        try:
            subprocess.run(
                command,
                check=True,
                timeout=launch_timeout,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return
        except Exception:
            continue
    raise RuntimeError(f"Could not open {xlsx_path} in Excel")


def save_workbook(*, timeout: float | None = 30.0) -> None:
    """Save the active Excel workbook via Cmd-S (overwrite in place)."""
    script = """
tell application "System Events"
    tell process "Microsoft Excel"
        set frontmost to true
        delay 0.2
        keystroke "s" using {command down}
        delay 0.5
    end tell
end tell
"""
    osascript(script, timeout=timeout)


def close_workbook(*, timeout: float | None = 30.0) -> None:
    """Close the active Excel workbook (no save prompt)."""
    script = """
tell application "Microsoft Excel"
    if (count workbooks) > 0 then
        close active workbook saving no
    end if
end tell
"""
    osascript(script, timeout=timeout)


def close_workbook_saving(*, timeout: float | None = 30.0) -> None:
    """Close the active workbook, persisting via the close handler.

    Excel writes the modified workbook back to its open path before
    tearing down the window. Use when the caller has staged a working
    copy and wants the post-Excel XML at that path. Mirrors the Word
    and PowerPoint oracle's close-with-save pattern.
    """
    script = f"""
tell application "{EXCEL_PROCESS_NAME}"
    if (count workbooks) > 0 then
        close active workbook saving yes
    end if
end tell
"""
    osascript(script, timeout=timeout)


def list_open_workbook_names() -> list[str]:
    """Return the names of workbooks currently open in Excel.

    Returns [] if Excel is not running, the AppleScript bridge times
    out, or the call fails. Notably catches `subprocess.TimeoutExpired`:
    sending `tell application "Microsoft Excel"` to a not-running Excel
    triggers a cold launch that frequently exceeds the 10s budget.

    Uses the `name of every workbook` idiom rather than
    `repeat with w in workbooks / name of w`. Empirically the
    latter hangs the AppleScript engine on Office for Mac M365 16.x
    even when the former returns instantly with the same underlying
    data — see Spec 022 baseline-collection notes.
    """
    script = f"""
tell application "{EXCEL_PROCESS_NAME}"
    return name of every workbook
end tell
"""
    try:
        out = osascript(script, timeout=10.0)
    except (RuntimeError, subprocess.TimeoutExpired):
        return []
    if not out:
        return []
    return [token.strip() for token in out.split(",") if token.strip()]


def is_workbook_open(path: Path) -> bool:
    """True if a workbook matching `path.name` is currently open."""
    return path.name in list_open_workbook_names()


def find_repair_dialog_text(
    patterns: tuple[str, ...] = REPAIR_DIALOG_PATTERNS,
) -> str | None:
    """If a modal alert is currently presented by Excel and its body
    text matches any of `patterns` (case-insensitive substring),
    return the full text. Otherwise None.

    Uses System Events UI scripting; relies on Excel presenting the
    alert as a child sheet/window of its main frame. Same shape as
    the PowerPoint and Word oracle scanners.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{EXCEL_PROCESS_NAME}") then return ""
    tell process "{EXCEL_PROCESS_NAME}"
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
    presented by Excel. Returns True if the button was found and
    clicked. Mirrors `pptx.osa.click_dialog_button` /
    `tools/oracle/word_window.click_dialog_button`.

    Uses System Events UI scripting; the `repeat with X in
    COLLECTION` idiom is fine here — the hang documented in
    `list_open_workbook_names` is specific to `tell application
    "Microsoft Excel"`, not System Events.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{EXCEL_PROCESS_NAME}") then return "false"
    tell process "{EXCEL_PROCESS_NAME}"
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


def dismiss_any_leftover_modal(*, timeout: float | None = 5.0) -> bool:
    """Best-effort dismissal of any leftover modal sheet via Escape.

    Excel sometimes shows a *secondary* info modal after the primary
    repair dialog and after close-with-save returns — wording like
    "Excel was able to open the file by repairing or removing the
    unreadable content" with `View` / `Delete` / X-close buttons.
    Those button labels aren't in `REPAIR_DIALOG_ACCEPT_BUTTON_LABELS`
    (we don't want to "View" the log; we want to dismiss). Sending
    Escape to the frontmost Excel window discards the modal in one
    shot.

    Returns True if a key event was dispatched (best-effort signal —
    the function can't reliably tell whether the modal was actually
    closed). Use as a finalize-after-close cleanup.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{EXCEL_PROCESS_NAME}") then return "false"
    tell process "{EXCEL_PROCESS_NAME}"
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


def dismiss_repair_dialog(
    *,
    accept_labels: tuple[str, ...] = REPAIR_DIALOG_ACCEPT_BUTTON_LABELS,
) -> tuple[bool, str | None, str | None]:
    """Detect and click through Excel's repair dialog.

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
