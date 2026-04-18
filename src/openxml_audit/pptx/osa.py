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
    "get_front_window_id",
    "get_slideshow_window_id",
    "launch_powerpoint",
    "open_presentation_via_ui",
]


POWERPOINT_APP_ID = "com.microsoft.Powerpoint"
POWERPOINT_PROCESS_NAME = "Microsoft PowerPoint"
POWERPOINT_APP_BUNDLE = Path("/Applications/Microsoft PowerPoint.app")


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
