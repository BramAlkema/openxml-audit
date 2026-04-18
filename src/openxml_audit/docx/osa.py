"""Word-specific AppleScript / JXA helpers for UI-driven opens and saves.

Builds on `openxml_audit.osa` primitives. macOS + Microsoft Word only.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

from openxml_audit.osa import launch_app, osascript

__all__ = [
    "WORD_APP_BUNDLE",
    "WORD_APP_ID",
    "WORD_PROCESS_NAME",
    "close_document",
    "launch_word",
    "open_document",
    "save_document",
]


WORD_APP_ID = "com.microsoft.Word"
WORD_PROCESS_NAME = "Microsoft Word"
WORD_APP_BUNDLE = Path("/Applications/Microsoft Word.app")


def launch_word() -> None:
    """Launch Word.app (via `open -W`)."""
    launch_app(WORD_APP_BUNDLE, WORD_PROCESS_NAME)


def open_document(
    docx_path: Path,
    *,
    timeout: float | None = 30.0,
) -> None:
    """Open a `.docx` in Word via `open -b` / `open -a`."""
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
            return
        except Exception:
            continue
    raise RuntimeError(f"Could not open {docx_path} in Word")


def save_document(*, timeout: float | None = 30.0) -> None:
    """Save the active Word document via Cmd-S (overwrite in place)."""
    script = """
tell application "System Events"
    tell process "Microsoft Word"
        set frontmost to true
        delay 0.2
        keystroke "s" using {command down}
        delay 0.5
    end tell
end tell
"""
    osascript(script, timeout=timeout)


def close_document(*, timeout: float | None = 30.0) -> None:
    """Close the active Word document (Cmd-W, no save prompt)."""
    script = """
tell application "Microsoft Word"
    if (count documents) > 0 then
        close active document saving no
    end if
end tell
"""
    osascript(script, timeout=timeout)
