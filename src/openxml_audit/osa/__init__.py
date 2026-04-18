"""macOS AppleScript / JXA primitives for driving Office apps.

Format-neutral building blocks used by format-specific oracle-authoring
tooling: `openxml_audit.pptx.osa` for PowerPoint today, and future
`docx.osa` / `xlsx.osa` layers for Word and Excel.

macOS only — requires `osascript` and the target app installed. Imports
succeed on any platform; calls raise at runtime on non-macOS.

Lifted from `svg2ooxml/tools/visual/pptx_window.py`; the sibling repo's
tune-loop screen-capture machinery stays there (ADR-002: converter-side
validation belongs in converter repos).
"""

from __future__ import annotations

import shlex
import subprocess
from pathlib import Path

__all__ = [
    "applescript_quote",
    "get_window_id_via_jxa",
    "launch_app",
    "osascript",
    "osascript_jxa",
]


def applescript_quote(value: str) -> str:
    """Escape a Python string for safe inclusion in an AppleScript string literal."""
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def osascript(script: str, *, timeout: float | None = 30.0) -> str:
    """Run an AppleScript via `osascript` and return stripped stdout."""
    result = subprocess.run(
        ["osascript", "-"],
        input=script,
        text=True,
        capture_output=True,
        check=False,
        timeout=timeout,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "osascript failed")
    return result.stdout.strip()


def osascript_jxa(script: str, *, timeout: float | None = 30.0) -> str:
    """Run a JavaScript-for-Automation script via `osascript -l JavaScript`."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-"],
        input=script,
        text=True,
        capture_output=True,
        check=False,
        timeout=timeout,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "osascript failed")
    return result.stdout.strip()


def launch_app(
    app_bundle_path: str | Path | None,
    process_name: str,
) -> None:
    """Launch a macOS app by bundle path (preferred) or process name fallback."""
    target: str
    if app_bundle_path is not None and Path(app_bundle_path).exists():
        target = str(app_bundle_path)
    else:
        target = process_name
    launch_cmd = f"open -W -a {shlex.quote(target)}"
    subprocess.Popen(
        ["/bin/zsh", "-lc", launch_cmd],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )


def get_window_id_via_jxa(
    owner_names: tuple[str, ...],
    *,
    name_contains: tuple[str, ...] = (),
    name_excludes: tuple[str, ...] = (),
) -> str:
    """Return the largest on-screen window ID matching owner names and name filters.

    Returns the empty string if no match is found. Uses CGWindowListCopyWindowInfo
    via JXA, with fallback from on-screen-only to all windows.
    """
    owners = ", ".join(f'"{owner.lower()}"' for owner in owner_names)
    name_contains_list = ", ".join(f'"{name.lower()}"' for name in name_contains)
    name_excludes_list = ", ".join(f'"{name.lower()}"' for name in name_excludes)
    script = f"""
ObjC.import("CoreGraphics");
var owners = [{owners}];
var nameContains = [{name_contains_list}];
var nameExcludes = [{name_excludes_list}];
function ownerMatches(w) {{
    var owner = (w.kCGWindowOwnerName || "").toLowerCase();
    return owners.indexOf(owner) !== -1;
}}
function nameMatches(w) {{
    if (!nameContains.length && !nameExcludes.length) {{
        return true;
    }}
    var name = (w.kCGWindowName || "").toLowerCase();
    if (nameContains.length) {{
        var ok = false;
        for (var i = 0; i < nameContains.length; i++) {{
            if (name.indexOf(nameContains[i]) !== -1) {{
                ok = true;
                break;
            }}
        }}
        if (!ok) {{
            return false;
        }}
    }}
    for (var j = 0; j < nameExcludes.length; j++) {{
        if (name.indexOf(nameExcludes[j]) !== -1) {{
            return false;
        }}
    }}
    return true;
}}
function area(w) {{
    var b = w.kCGWindowBounds || {{}};
    return (b.Width || 0) * (b.Height || 0);
}}
function findMatches(windows, requireOnscreen) {{
    if (!windows || !windows.filter) {{
        return [];
    }}
    return windows.filter(function (w) {{
        if (!ownerMatches(w) || !nameMatches(w)) {{
            return false;
        }}
        if (w.kCGWindowLayer !== 0) {{
            return false;
        }}
        if (requireOnscreen && !w.kCGWindowIsOnscreen) {{
            return false;
        }}
        return true;
    }});
}}
var list = $.CGWindowListCopyWindowInfo($.kCGWindowListOptionOnScreenOnly, $.kCGNullWindowID);
var windows = ObjC.deepUnwrap(ObjC.castRefToObject(list));
if (!windows || !windows.filter) {{
    windows = [];
}}
var matches = findMatches(windows, true);
if (!matches.length) {{
    matches = findMatches(windows, false);
}}
if (!matches.length) {{
    var listAll = $.CGWindowListCopyWindowInfo($.kCGWindowListOptionAll, $.kCGNullWindowID);
    var windowsAll = ObjC.deepUnwrap(ObjC.castRefToObject(listAll));
    if (!windowsAll || !windowsAll.filter) {{
        windowsAll = [];
    }}
    matches = findMatches(windowsAll, false);
}}
if (!matches.length) {{
    "";
}} else {{
    matches.sort(function (a, b) {{ return area(b) - area(a); }});
    matches[0].kCGWindowNumber.toString();
}}
"""
    return osascript_jxa(script)
