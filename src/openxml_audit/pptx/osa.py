"""PowerPoint-specific AppleScript / JXA helpers for window control and UI-driven opens.

Builds on `openxml_audit.osa` primitives. macOS + PowerPoint only.
"""

from __future__ import annotations

import subprocess
import time
from pathlib import Path

from openxml_audit.osa import (
    applescript_quote,
    get_window_id_via_jxa,
    launch_app,
    osascript,
)

__all__ = [
    "POWERPOINT_APP_BUNDLE",
    "POWERPOINT_APP_ID",
    "POWERPOINT_PROCESS_NAME",
    "REPAIR_DIALOG_ACCEPT_BUTTON_LABELS",
    "REPAIR_DIALOG_PATTERNS",
    "click_dialog_button",
    "close_presentation",
    "close_presentation_saving",
    "dismiss_any_leftover_modal",
    "dismiss_repair_dialog",
    "find_repair_dialog_text",
    "get_front_window_id",
    "get_slideshow_window_id",
    "is_presentation_open",
    "launch_powerpoint",
    "list_open_presentation_names",
    "open_presentation_via_ui",
    "save_presentation",
]


POWERPOINT_APP_ID = "com.microsoft.Powerpoint"
POWERPOINT_PROCESS_NAME = "Microsoft PowerPoint"
POWERPOINT_APP_BUNDLE = Path("/Applications/Microsoft PowerPoint.app")


# Repair-dialog text patterns. PowerPoint surfaces a "found a problem"
# / "unreadable content" modal when it repairs a file on open. Generous
# match list — different builds and different repair categories
# produce different exact strings.
REPAIR_DIALOG_PATTERNS: tuple[str, ...] = (
    "found a problem",
    "unreadable content",
    "repair",
    "recover",
    "could not be opened",
    "experienced an error",
    "error trying to open",
    "errors were detected",
)


# Button labels to try (in order) when accepting PowerPoint's repair
# dialog. PowerPoint's modals across Office for Mac M365 have used
# Yes / Repair / OK / Recover variations across versions, and the
# "would you like to try to repair?" sheet typically exposes a
# `Repair` primary button. Order matters: try the most-specific
# affirmative first, fall back to generic OK.
REPAIR_DIALOG_ACCEPT_BUTTON_LABELS: tuple[str, ...] = (
    "Repair",
    "Yes",
    "Recover",
    "OK",
    "Open",
)


def launch_powerpoint() -> None:
    """Launch PowerPoint.app (via `open -W`)."""
    launch_app(POWERPOINT_APP_BUNDLE, POWERPOINT_PROCESS_NAME)


def open_presentation_via_ui(
    pptx_path: Path,
    *,
    timeout: float | None = 30.0,
) -> None:
    """Open a .pptx in PowerPoint, falling back to scripted Go-To-Folder keystrokes."""
    pptx_path = pptx_path.resolve()
    launch_timeout = max(5.0, min(timeout or 30.0, 15.0))
    launch_attempts = (
        ["open", "-b", POWERPOINT_APP_ID, str(pptx_path)],
        ["open", "-a", POWERPOINT_PROCESS_NAME, str(pptx_path)],
    )
    for command in launch_attempts:
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

    target_posix = applescript_quote(str(pptx_path))
    launch_powerpoint()
    script = f"""
set targetPosix to {target_posix}
delay 1.0
tell application "System Events"
    keystroke "o" using {{command down}}
    delay 0.5
    keystroke "g" using {{command down, shift down}}
    delay 0.3
    keystroke targetPosix
    delay 0.2
    key code 36
    delay 0.4
    key code 36
end tell
"""
    osascript(script, timeout=timeout)


def get_front_window_id(pptx_path: Path, delay: float) -> str:
    """Return PowerPoint's front-window ID after opening the given .pptx."""
    pptx_path = pptx_path.resolve()
    try:
        open_presentation_via_ui(pptx_path, timeout=max(10.0, delay + 10.0))
        script = f"""
delay {delay}
tell application "System Events"
    tell process "{POWERPOINT_PROCESS_NAME}"
        set frontmost to true
        delay 0.2
        set winId to value of attribute "AXWindowNumber" of front window
    end tell
end tell
return winId
"""
        win_id = osascript(script)
    except RuntimeError:
        win_id = ""
    if not win_id:
        win_id = get_window_id_via_jxa((POWERPOINT_PROCESS_NAME,))
    return win_id


def list_open_presentation_names() -> list[str]:
    """Return the names of presentations currently open in PowerPoint.

    Returns [] if PowerPoint is not running, the AppleScript bridge
    times out, or the call fails. Notably catches
    `subprocess.TimeoutExpired`: sending `tell application "Microsoft
    PowerPoint"` to a not-running PowerPoint triggers a cold launch
    that frequently exceeds the 10s budget.

    Uses the `name of every presentation` idiom rather than
    `repeat with p in presentations / name of p`. Empirically the
    latter hangs PowerPoint's AppleScript engine on Office for Mac
    M365 16.x even when the former returns instantly with the same
    underlying data — see Spec 022 baseline-collection notes.
    """
    script = f"""
tell application "{POWERPOINT_PROCESS_NAME}"
    return name of every presentation
end tell
"""
    try:
        out = osascript(script, timeout=10.0)
    except (RuntimeError, subprocess.TimeoutExpired):
        return []
    if not out:
        return []
    return [token.strip() for token in out.split(",") if token.strip()]


def is_presentation_open(path: Path) -> bool:
    """True if a presentation matching `path.name` is currently open."""
    return path.name in list_open_presentation_names()


def save_presentation(*, timeout: float | None = 30.0) -> None:
    """Save the active PowerPoint presentation via Cmd-S (overwrite in place).

    Mirrors `docx.osa.save_document`: PowerPoint's AppleScript `save`
    on a presentation is sometimes flaky between builds; routing
    through System Events keystrokes is the proven path.
    """
    script = f"""
tell application "System Events"
    tell process "{POWERPOINT_PROCESS_NAME}"
        set frontmost to true
        delay 0.2
        keystroke "s" using {{command down}}
        delay 0.5
    end tell
end tell
"""
    osascript(script, timeout=timeout)


def close_presentation(*, timeout: float | None = 30.0) -> None:
    """Close the active presentation without saving (mirror docx.osa.close_document)."""
    script = f"""
tell application "{POWERPOINT_PROCESS_NAME}"
    if (count presentations) > 0 then
        close active presentation saving no
    end if
end tell
"""
    osascript(script, timeout=timeout)


def close_presentation_saving(*, timeout: float | None = 30.0) -> None:
    """Close the active presentation, persisting via the close handler.

    PowerPoint writes the modified presentation back to its open path
    before tearing down the window. Use when the caller has staged a
    working copy and wants the post-PowerPoint XML at that path.
    """
    script = f"""
tell application "{POWERPOINT_PROCESS_NAME}"
    if (count presentations) > 0 then
        close active presentation saving yes
    end if
end tell
"""
    osascript(script, timeout=timeout)


def find_repair_dialog_text(
    patterns: tuple[str, ...] = REPAIR_DIALOG_PATTERNS,
) -> str | None:
    """If a modal alert is currently presented by PowerPoint and its
    body text matches any of `patterns` (case-insensitive substring),
    return the full text. Otherwise None.

    Uses System Events UI scripting; relies on PowerPoint presenting
    the alert as a child sheet/window of its main frame. Tested
    empirically; subject to PowerPoint UI changes.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{POWERPOINT_PROCESS_NAME}") then return ""
    tell process "{POWERPOINT_PROCESS_NAME}"
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


def get_slideshow_window_id(timeout: float = 5.0) -> str:
    """Poll for the PowerPoint slideshow window ID, excluding the presenter view."""
    end_time = time.time() + timeout
    while time.time() < end_time:
        for owner in ("Microsoft PowerPoint Slide Show", "PowerPoint Slide Show"):
            win_id = get_window_id_via_jxa((owner,), name_excludes=("presenter",))
            if win_id:
                return win_id
        win_id = get_window_id_via_jxa(
            (POWERPOINT_PROCESS_NAME,),
            name_contains=("slide show", "slideshow"),
            name_excludes=("presenter",),
        )
        if win_id:
            return win_id
        time.sleep(0.2)
    return ""


def click_dialog_button(button_label: str) -> bool:
    """Click a button by exact label in any sheet/dialog currently
    presented by PowerPoint. Returns True if the button was found and
    clicked. Mirrors `tools/oracle/word_window.click_dialog_button`.

    Uses System Events UI scripting; iterates over `windows` and
    `sheets of windows` of the PowerPoint process. The
    `repeat with X in COLLECTION` idiom is fine here — the hang
    documented in `list_open_presentation_names` is specific to
    `tell application "Microsoft PowerPoint"`, not System Events.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{POWERPOINT_PROCESS_NAME}") then return "false"
    tell process "{POWERPOINT_PROCESS_NAME}"
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

    PowerPoint may show a follow-up info modal after the primary
    repair dialog or after close-with-save returns. Button labels for
    these vary; sending Escape to the frontmost PowerPoint window
    discards the modal in one shot.
    """
    script = f"""
tell application "System Events"
    if not (exists process "{POWERPOINT_PROCESS_NAME}") then return "false"
    tell process "{POWERPOINT_PROCESS_NAME}"
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
    """Detect and click through PowerPoint's repair dialog.

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
