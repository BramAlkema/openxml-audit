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
    "close_workbook",
    "launch_excel",
    "open_workbook",
    "save_workbook",
]


EXCEL_APP_ID = "com.microsoft.Excel"
EXCEL_PROCESS_NAME = "Microsoft Excel"
EXCEL_APP_BUNDLE = Path("/Applications/Microsoft Excel.app")


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
