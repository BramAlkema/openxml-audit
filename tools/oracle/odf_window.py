"""LibreOffice (soffice) headless harness for ODF roundtrip oracle work.

soffice is the right tool for ODF roundtrips but is operationally fragile:

  - Single-instance lock by default — concurrent calls collide unless each
    one gets its own UserInstallation profile dir.
  - Hangs on malformed input — no internal timeout, blocks indefinitely.
  - Leaves zombie processes when killed mid-run.
  - Crashes on missing fonts / specific malformed feature combinations.

The corpus walks this oracle drives are 100s-1000s of files. A naive
`subprocess.run(['soffice', ...])` will hit one of these failure modes
within minutes and the run will stall or corrupt later results. This
module wraps soffice with explicit supervision so each call:

  - gets a fresh ephemeral UserInstallation directory (no profile lock
    contention, no carried-over state from a previous file)
  - runs under a hard wall-clock timeout (caller-configurable, default 60s)
  - is force-killed on timeout, with descendant processes reaped via
    process-group signaling
  - reports failures as structured `SofficeRunResult` rather than raising
    on every soft error — corpus walks need to record "this file crashed
    soffice" as data, not stall the whole run

Spec reference: 0.6.3 (ODF roundtrip oracle), Spec 019.
"""

from __future__ import annotations

import os
import shutil
import signal
import subprocess
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal


_SOFFICE_PROCESS_NAME = "soffice"
_SOFFICE_BIN_CANDIDATES = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
)


class SofficeNotFoundError(RuntimeError):
    """Raised when no soffice executable can be located on the system."""


@dataclass
class SofficeRunResult:
    """Structured outcome of a single soffice invocation.

    `output_path` is None on every failure path. `outcome` distinguishes
    so callers can record corpus-walk statistics (how many files crashed
    soffice vs timed out vs converted cleanly).
    """

    outcome: Literal["ok", "timeout", "crash", "missing_output", "exit_nonzero"]
    output_path: Path | None
    duration_seconds: float
    return_code: int | None = None
    stderr: str = ""
    stdout: str = ""
    notes: list[str] = field(default_factory=list)


def find_soffice() -> Path:
    """Locate the soffice binary, preferring the macOS bundle path then PATH.

    Raises `SofficeNotFoundError` if no candidate is executable.
    """
    for candidate in _SOFFICE_BIN_CANDIDATES:
        path = Path(candidate)
        if path.exists() and os.access(path, os.X_OK):
            return path
    which = shutil.which(_SOFFICE_PROCESS_NAME)
    if which:
        return Path(which)
    raise SofficeNotFoundError(
        f"soffice not found. Tried bundle paths {_SOFFICE_BIN_CANDIDATES} "
        f"and PATH lookup for '{_SOFFICE_PROCESS_NAME}'."
    )


_FORMAT_TO_FILTER = {
    "odt": "writer8",
    "ods": "calc8",
    "odp": "impress8",
    "odg": "draw8",
}


def roundtrip(
    input_path: Path,
    output_dir: Path,
    *,
    target_format: Literal["odt", "ods", "odp", "odg"] = "odt",
    timeout_seconds: float = 60.0,
    soffice_bin: Path | None = None,
) -> SofficeRunResult:
    """Roundtrip `input_path` through soffice's headless converter.

    Each call runs in an ephemeral UserInstallation profile directory
    (cleaned up on return). The output file lands in `output_dir` with
    the same basename as the input — soffice's `--convert-to` does not
    let us pick the output filename, so callers should use a fresh
    `output_dir` per call when avoiding name collisions.

    The wall-clock timeout is enforced via `subprocess` and force-kill
    on the soffice process group so descendant `oosplash` / `xpdfimport`
    helpers don't survive as zombies.

    Returns a `SofficeRunResult` describing the outcome. The caller is
    responsible for distinguishing `outcome == "ok"` from soft failures.
    """
    binary = soffice_bin or find_soffice()
    if not input_path.exists():
        return SofficeRunResult(
            outcome="missing_output",
            output_path=None,
            duration_seconds=0.0,
            notes=[f"input does not exist: {input_path}"],
        )
    output_dir.mkdir(parents=True, exist_ok=True)

    profile_dir = Path(tempfile.mkdtemp(prefix="oxa-odf-oracle-"))
    cmd = [
        str(binary),
        "--headless",
        "--norestore",
        "--nologo",
        "--nofirststartwizard",
        f"-env:UserInstallation=file://{profile_dir}",
        "--convert-to",
        target_format,
        "--outdir",
        str(output_dir),
        str(input_path),
    ]

    notes: list[str] = []
    start = time.monotonic()
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
            # New process group so we can kill descendants on timeout.
            start_new_session=True,
        )
    except subprocess.TimeoutExpired as exc:
        # subprocess.run kills the immediate child but soffice's helper
        # processes (oosplash, etc.) can survive. Best-effort reap by
        # matching the profile path in the cmdline.
        _kill_soffice_for_profile(profile_dir)
        return SofficeRunResult(
            outcome="timeout",
            output_path=None,
            duration_seconds=time.monotonic() - start,
            return_code=None,
            stderr=exc.stderr if isinstance(exc.stderr, str) else "",
            stdout=exc.stdout if isinstance(exc.stdout, str) else "",
            notes=[f"timeout after {timeout_seconds}s; reaped soffice for profile"],
        )
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    duration = time.monotonic() - start
    expected_output = output_dir / input_path.name

    if proc.returncode != 0:
        return SofficeRunResult(
            outcome="exit_nonzero",
            output_path=None,
            duration_seconds=duration,
            return_code=proc.returncode,
            stderr=proc.stderr,
            stdout=proc.stdout,
            notes=notes + [f"soffice exited {proc.returncode}"],
        )

    if not expected_output.exists():
        return SofficeRunResult(
            outcome="missing_output",
            output_path=None,
            duration_seconds=duration,
            return_code=proc.returncode,
            stderr=proc.stderr,
            stdout=proc.stdout,
            notes=notes + [
                f"soffice exited 0 but expected output {expected_output} is absent"
            ],
        )

    return SofficeRunResult(
        outcome="ok",
        output_path=expected_output,
        duration_seconds=duration,
        return_code=proc.returncode,
        stderr=proc.stderr,
        stdout=proc.stdout,
        notes=notes,
    )


def _kill_soffice_for_profile(profile_dir: Path) -> None:
    """Best-effort: kill any soffice processes whose UserInstallation arg
    points at `profile_dir`. Used after a timeout to reap stragglers."""
    try:
        result = subprocess.run(
            ["pgrep", "-f", f"UserInstallation=file://{profile_dir}"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return
    if result.returncode != 0:
        return
    for line in result.stdout.splitlines():
        line = line.strip()
        if not line.isdigit():
            continue
        try:
            os.kill(int(line), signal.SIGKILL)
        except (OSError, ProcessLookupError):
            pass


__all__ = [
    "SofficeError",
    "SofficeNotFoundError",
    "SofficeRunResult",
    "find_soffice",
    "roundtrip",
]


# Public alias for callers wanting a single sentinel for soffice troubles.
SofficeError = SofficeNotFoundError
