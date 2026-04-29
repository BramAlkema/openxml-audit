"""Preflight environment checks for roundtrip oracles.

Surfaces missing prerequisites with actionable error messages before the
engine attempts a roundtrip. Run with:

  python -m tools.oracle.preflight                # check all four engines
  python -m tools.oracle.preflight --engine word
  python -m tools.oracle.preflight --engine excel
  python -m tools.oracle.preflight --engine powerpoint
  python -m tools.oracle.preflight --engine odf

or import `check_word()`, `check_excel()`, `check_powerpoint()`,
`check_odf()` programmatically — the pytest skip hooks for each
roundtrip oracle module call these.

A clean preflight means the engine has a fair chance of running.
macOS Privacy & Security permission grants are gated by user
interaction and cannot be inspected without prompting, so the first
real roundtrip after a clean preflight may still trigger a permission
prompt for Automation (Office apps) or Accessibility (System Events
keystrokes).

See `docs/oracle_permissions.md` for the macOS setup checklist.
"""

from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

# Existing constant — preserved for backward-compat with imports of
# `tools.oracle.preflight.WORD_APP_BUNDLE` from the Word oracle code.
WORD_APP_BUNDLE = Path("/Applications/Microsoft Word.app")
EXCEL_APP_BUNDLE = Path("/Applications/Microsoft Excel.app")
POWERPOINT_APP_BUNDLE = Path("/Applications/Microsoft PowerPoint.app")

_SOFFICE_BIN_CANDIDATES = (
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    Path("/usr/bin/soffice"),
    Path("/usr/local/bin/soffice"),
    Path("/opt/homebrew/bin/soffice"),
)


@dataclass
class PreflightStatus:
    ok: bool
    issues: list[str]
    engine: str = ""


def _check_office_app(
    *,
    engine: str,
    process_name: str,
    bundle: Path,
    version_script: str,
) -> PreflightStatus:
    """Common shape: macOS + osascript + app installed + AppleScript-reachable."""
    issues: list[str] = []

    if platform.system() != "Darwin":
        issues.append(
            f"Not running on macOS — the {engine} roundtrip oracle is macOS-only."
        )
        return PreflightStatus(ok=False, issues=issues, engine=engine)

    if shutil.which("osascript") is None:
        issues.append("`osascript` not found on PATH — required for AppleScript dispatch.")

    if not bundle.exists():
        issues.append(
            f"{process_name} not found at {bundle}. Install it from Microsoft 365 "
            "(or skip this engine)."
        )

    # Best-effort liveness check: dispatch a trivial AppleScript. On
    # first run with no Automation permission this triggers Apple's
    # consent prompt; on subsequent runs without permission the
    # AppleScript fails with a -1743 error.
    if shutil.which("osascript") is not None and bundle.exists():
        try:
            result = subprocess.run(
                ["osascript", "-e", version_script],
                capture_output=True,
                text=True,
                timeout=15.0,
            )
            if result.returncode != 0:
                stderr = result.stderr.strip()
                issues.append(
                    f"Could not query {process_name}'s version via AppleScript. "
                    "Likely needs Automation permissions in System Settings → "
                    "Privacy & Security → Automation (grant your terminal app "
                    f"control over '{process_name}'). osascript stderr: "
                    f"{stderr or '(empty)'}"
                )
        except subprocess.TimeoutExpired:
            issues.append(
                f"AppleScript dispatch to {process_name} timed out (15s). "
                "The app may be cold-launching for the first time, or a "
                "permission prompt may be waiting on screen. Re-run the "
                "preflight after granting permission, or force-quit the app "
                "and retry."
            )

    return PreflightStatus(ok=not issues, issues=issues, engine=engine)


def check_word() -> PreflightStatus:
    return _check_office_app(
        engine="word",
        process_name="Microsoft Word",
        bundle=WORD_APP_BUNDLE,
        version_script='tell application "Microsoft Word" to return version',
    )


def check_excel() -> PreflightStatus:
    return _check_office_app(
        engine="excel",
        process_name="Microsoft Excel",
        bundle=EXCEL_APP_BUNDLE,
        version_script='tell application "Microsoft Excel" to return version',
    )


def check_powerpoint() -> PreflightStatus:
    return _check_office_app(
        engine="powerpoint",
        process_name="Microsoft PowerPoint",
        bundle=POWERPOINT_APP_BUNDLE,
        version_script='tell application "Microsoft PowerPoint" to return version',
    )


def check_odf() -> PreflightStatus:
    """ODF preflight: just verifies a soffice binary is on disk. The ODF
    oracle uses headless `soffice --convert-to`, not AppleScript, so no
    Privacy & Security permissions are involved."""
    issues: list[str] = []

    soffice = next((p for p in _SOFFICE_BIN_CANDIDATES if p.exists()), None)
    if soffice is None:
        which = shutil.which("soffice")
        if which:
            soffice = Path(which)
    if soffice is None:
        issues.append(
            "`soffice` not found. Install LibreOffice "
            f"(checked {[str(p) for p in _SOFFICE_BIN_CANDIDATES]} and PATH)."
        )
    return PreflightStatus(ok=not issues, issues=issues, engine="odf")


# Back-compat alias — the original Word-only API. Existing callers and
# tests keep working.
def check() -> PreflightStatus:
    """Backward-compatible alias for `check_word()`."""
    return check_word()


_ALL_ENGINES = {
    "word": check_word,
    "excel": check_excel,
    "powerpoint": check_powerpoint,
    "odf": check_odf,
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--engine",
        choices=sorted(_ALL_ENGINES),
        default=None,
        help="Check a single engine. Default: check all four.",
    )
    args = parser.parse_args()

    if args.engine is None:
        engines = list(_ALL_ENGINES)
    else:
        engines = [args.engine]

    overall_ok = True
    for engine in engines:
        status = _ALL_ENGINES[engine]()
        if status.ok:
            print(f"  {engine:10s} OK")
        else:
            overall_ok = False
            print(f"  {engine:10s} FAIL", file=sys.stderr)
            for issue in status.issues:
                print(f"    - {issue}", file=sys.stderr)

    if not overall_ok:
        print(
            "\nSee docs/oracle_permissions.md for the macOS setup checklist.",
            file=sys.stderr,
        )
    return 0 if overall_ok else 1


if __name__ == "__main__":
    sys.exit(main())
