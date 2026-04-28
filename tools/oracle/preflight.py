"""Preflight environment check for the Word roundtrip oracle.

Surfaces missing prerequisites with actionable error messages before the
engine attempts a roundtrip. Run as `python -m tools.oracle.preflight`
or import `check()` programmatically — the latter is what the
`requires_word_app` pytest skip hook calls.
"""

from __future__ import annotations

import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

WORD_APP_BUNDLE = Path("/Applications/Microsoft Word.app")


@dataclass
class PreflightStatus:
    ok: bool
    issues: list[str]


def check() -> PreflightStatus:
    """Return a structured report. `ok=True` means the engine has a fair
    chance of running. Permission grants in System Settings cannot be
    inspected without prompting, so a clean preflight does not guarantee
    the first run will succeed."""
    issues: list[str] = []

    if platform.system() != "Darwin":
        issues.append(
            "Not running on macOS — the Word roundtrip oracle is macOS-only."
        )
        return PreflightStatus(ok=False, issues=issues)

    if shutil.which("osascript") is None:
        issues.append("`osascript` not found on PATH — required for AppleScript dispatch.")

    if not WORD_APP_BUNDLE.exists():
        issues.append(
            f"Microsoft Word not found at {WORD_APP_BUNDLE}. Install Word for Mac "
            "(Microsoft 365) and re-run."
        )

    # A best-effort liveness check: dispatch a trivial AppleScript that
    # asks for Word's version. This will trigger Apple's permission
    # prompts on first run if UI scripting isn't allowed yet.
    if shutil.which("osascript") is not None and WORD_APP_BUNDLE.exists():
        try:
            result = subprocess.run(
                ["osascript", "-e", 'tell application "Microsoft Word" to return version'],
                capture_output=True,
                text=True,
                timeout=10.0,
            )
            if result.returncode != 0:
                stderr = result.stderr.strip()
                issues.append(
                    "Could not query Word's version via AppleScript. Likely needs "
                    "Automation permissions in System Settings → Privacy & Security → "
                    f"Automation. osascript stderr: {stderr or '(empty)'}"
                )
        except subprocess.TimeoutExpired:
            issues.append(
                "AppleScript dispatch to Word timed out. Word may be unresponsive; "
                "force-quit it and retry."
            )

    return PreflightStatus(ok=not issues, issues=issues)


def main() -> int:
    status = check()
    if status.ok:
        print("Word roundtrip oracle preflight: OK")
        return 0
    print("Word roundtrip oracle preflight: FAIL", file=sys.stderr)
    for issue in status.issues:
        print(f"  - {issue}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
