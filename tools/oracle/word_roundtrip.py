"""Roundtrip a DOCX through Microsoft Word for Mac and capture the post-Word file.

The diff between the input and the post-Word file is the oracle for any
"would Word repair this?" question. See `specs/011-word-roundtrip-oracle.md`.

This module is developer-machine infrastructure: macOS + Microsoft Word
required. None of it ships in the importable `openxml_audit` package.

Status: Phase 1 build. The end-to-end flow has not been smoke-tested
against a real Word installation in this commit; the AppleScript
commands are coded against Word for Mac M365's documented surface and
may need empirical adjustment on first run. See the spec's open
questions.
"""

from __future__ import annotations

import os
import shutil
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from tools.oracle import word_window

# Word for Mac is App Sandboxed; by default it can only access files
# under the user's Documents/Desktop/Downloads, iCloud Drive, and its
# Office Group Container. Staging in /tmp (mkdtemp default) silently
# fails: the AppleScript `open` hangs waiting for a sandbox-permission
# response that never arrives. We stage in a Documents-scoped subdirectory
# instead, which Word's sandbox accepts without prompting.
#
# Override via the WORD_ORACLE_STAGE env var if you've granted Word Full
# Disk Access and want a different location.
_DEFAULT_STAGE_PARENT = Path.home() / "Documents" / ".word_oracle_runs"


@dataclass
class RoundtripResult:
    """Structured outcome of one Word roundtrip."""

    input_path: Path
    output_path: Path
    repair_dialog_seen: bool
    repair_dialog_text: str | None
    word_version: str | None
    elapsed_seconds: float
    staged_input_path: Path = field(default_factory=lambda: Path())


class RoundtripError(RuntimeError):
    """Raised when the engine cannot complete a roundtrip (timeout, missing
    Word, refused permissions, etc.). Distinct from "Word ran but rejected
    the file" which is reported via `RoundtripResult.repair_dialog_seen`."""


def _stage_root() -> Path:
    """Resolve the staging parent directory. Defaults to a Documents-scoped
    subdirectory because Word's App Sandbox grants access there by default;
    override with WORD_ORACLE_STAGE for users who've granted Full Disk Access."""
    override = os.environ.get("WORD_ORACLE_STAGE")
    base = Path(override).expanduser() if override else _DEFAULT_STAGE_PARENT
    base.mkdir(parents=True, exist_ok=True)
    return base


def _stage_input(input_path: Path, output_dir: Path | None) -> tuple[Path, Path, Path]:
    """Stage the input in a private staging directory.

    Word's AppleScript surface lacks a working `save as` and a working
    `save` on its windows; the only reliable persist path is
    `close active document saving yes`, which writes back to the
    document's current location. So we open Word against a *working
    copy* and let Word overwrite that copy when it closes-with-save. The
    untouched original lives next to it for diffing.

    Returns (work_dir, working_copy_to_open, output_path). When
    `output_dir` is None, the working copy is also the output — they
    refer to the same file.
    """
    work_dir = _stage_root() / f"run-{uuid.uuid4().hex[:8]}"
    work_dir.mkdir(parents=True, exist_ok=False)

    # Reserve the original under a distinct name so the diff has both
    # ends.
    original_copy = work_dir / f"original_{input_path.name}"
    shutil.copy2(input_path, original_copy)

    working = work_dir / input_path.name
    shutil.copy2(input_path, working)

    if output_dir is not None:
        # Caller wants the output relocated. Resolve at the end of the
        # roundtrip; the working copy is what Word writes to in place.
        output_dir.mkdir(parents=True, exist_ok=True)
        output = output_dir / f"post_word_{input_path.name}"
    else:
        output = working

    return work_dir, working, output


def _wait_for_document_open(
    document_name: str, *, timeout: float, poll_interval: float = 0.5
) -> None:
    """Poll Word's `documents` collection until `document_name` appears or
    `timeout` expires. The check matches by display name (Word strips the
    path), so the staged-input filename should be unique among open docs."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        names = word_window.list_open_document_names()
        if any(name == document_name or name.endswith(f"/{document_name}") for name in names):
            return
        time.sleep(poll_interval)
    raise RoundtripError(
        f"Word did not open '{document_name}' within {timeout:.1f}s "
        f"(docs reachable: {word_window.list_open_document_names()})"
    )


def _try_dismiss_repair_dialog(
    *, accept: bool
) -> tuple[bool, str | None]:
    """Look for Word's repair dialog and dismiss it. Returns (was_seen, text)."""
    text = word_window.find_repair_dialog_text()
    if text is None:
        return False, None
    button = "Yes" if accept else "No"
    clicked = word_window.click_dialog_button(button)
    if not clicked:
        # Try alternate labels — some Word builds use "OK" / "Cancel"
        for alt in (("OK", "Cancel"), ("Recover", "Don't Recover")):
            label = alt[0] if accept else alt[1]
            if word_window.click_dialog_button(label):
                clicked = True
                break
    if not clicked:
        raise RoundtripError(
            f"Detected repair dialog but failed to dismiss "
            f"(target button: {button!r}, dialog text: {text!r})"
        )
    return True, text


def roundtrip(
    input_docx: Path,
    *,
    output_dir: Path | None = None,
    timeout: float = 60.0,
    accept_repair: bool = True,
) -> RoundtripResult:
    """Open `input_docx` in Word, save it through Word, return the saved path.

    Parameters
    ----------
    input_docx
        Path to the DOCX to roundtrip. Copied to a private temp directory
        before any Word operation; the source path is never modified.
    output_dir
        Directory to place the post-Word file. Defaults to the staging temp
        directory; the caller is responsible for cleanup either way (the
        full RoundtripResult exposes both paths).
    timeout
        Per-step timeout in seconds. The full roundtrip may take longer
        if Word is slow to launch.
    accept_repair
        When Word's "unreadable content" dialog appears, dismiss as Yes
        (preserve the post-repair XML) or No (cancel save and fail).

    Returns
    -------
    RoundtripResult
        Both file paths, the repair-dialog flag and text, the Word version
        string, and elapsed time. The caller diffs `input_path` vs
        `output_path` to derive the oracle verdict.
    """
    input_docx = input_docx.resolve()
    if not input_docx.exists():
        raise RoundtripError(f"Input file does not exist: {input_docx}")

    work_dir, staged, target = _stage_input(input_docx, output_dir)
    started = time.monotonic()

    word_window.launch_word_app()
    # Brief settling pause — `open -W` waits for Word to be reachable but
    # the AppleScript dispatcher can race the first `open` command.
    time.sleep(1.0)

    # Dispatch the open; if AppleScript hangs (Word's welcome screen,
    # sandbox prompt, etc.) the polling loop below still has a chance to
    # detect the document if Word eventually picks it up.
    word_window.open_document(staged, timeout=10.0)

    # Word may surface the repair dialog *before* the document name appears
    # in the documents list. Poll for either: the dialog, or the open
    # confirmation. Whichever arrives first is the next event we handle.
    repair_seen = False
    repair_text: str | None = None
    deadline = time.monotonic() + timeout
    seen_open = False
    while time.monotonic() < deadline:
        if not repair_seen:
            saw, text = _try_dismiss_repair_dialog(accept=accept_repair)
            if saw:
                repair_seen = True
                repair_text = text
                if not accept_repair:
                    raise RoundtripError(
                        f"Repair dialog seen and rejected by caller. "
                        f"Dialog text: {text!r}"
                    )
        if word_window.is_document_open(staged):
            seen_open = True
            break
        time.sleep(0.5)
    if not seen_open:
        raise RoundtripError(
            f"Word did not register '{staged.name}' as open within {timeout:.1f}s "
            f"(reachable docs: {word_window.list_open_document_names()})"
        )

    # Make sure the staged document is the active one before close-with-save
    word_window.activate_document(staged)

    # Persist via close-with-save: Word writes the modified document back
    # to its open path (the staged working copy). This is the only
    # AppleScript persist path that works in Word for Mac M365.
    try:
        word_window.close_active_document_saving()
    except word_window.OsascriptError as exc:
        raise RoundtripError(f"Close-with-save failed: {exc}") from exc

    elapsed = time.monotonic() - started

    if not staged.exists():
        raise RoundtripError(
            f"Close-with-save returned but the working copy is missing: {staged}"
        )

    # If the caller wanted the output relocated, copy the now-post-Word
    # working file to the requested output path.
    if target != staged:
        shutil.copy2(staged, target)

    return RoundtripResult(
        input_path=input_docx,
        output_path=target,
        staged_input_path=staged,
        repair_dialog_seen=repair_seen,
        repair_dialog_text=repair_text,
        word_version=word_window.word_version(),
        elapsed_seconds=elapsed,
    )
