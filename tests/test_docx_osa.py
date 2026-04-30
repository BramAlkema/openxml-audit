"""Tests for `openxml_audit.docx.osa` (Spec 030, 0.7.4).

The 0.7.4 consolidation moved Word window primitives from
`tools/oracle/word_window.py` (legacy, since Spec 011) into
`openxml_audit.docx.osa` (in-package, parallels `pptx.osa` /
`xlsx.osa`). `tools/oracle/word_window.py` is now a thin back-compat
shim that re-exports from the in-package layer.

These tests lock:

  - the symbols the in-package layer must expose
  - the back-compat shim re-exports those same symbols
  - the dismiss-helper and dialog-pattern shape contract
  - `OsascriptError` is still importable from the legacy path and
    still catches what existing `except` clauses caught
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from openxml_audit.docx import osa as docx_osa  # noqa: E402
from oracle import word_window  # noqa: E402  (back-compat shim)


# Symbols the in-package layer must expose. If anything here breaks,
# the consumers in tools/oracle/word_*.py break too.
_REQUIRED_OSA_SYMBOLS = {
    # Constants / module data
    "WORD_APP_BUNDLE",
    "WORD_APP_ID",
    "WORD_PROCESS_NAME",
    "REPAIR_DIALOG_PATTERNS",
    "REPAIR_DIALOG_ACCEPT_BUTTON_LABELS",
    # App lifecycle
    "launch_word",
    "launch_word_app",  # alias for launch_word
    "is_word_running",
    "word_version",
    # Document open/list/identify
    "open_document",
    "list_open_document_names",
    "is_document_open",
    "activate_document",
    # Save/close primitives
    "save_document",
    "close_document",
    "close_active_document",  # alias for close_document
    "close_document_saving",
    "close_active_document_saving",  # alias for close_document_saving
    # Dialog handling
    "find_repair_dialog_text",
    "click_dialog_button",
    "dismiss_repair_dialog",
    "dismiss_any_leftover_modal",
}


def test_docx_osa_exports_full_word_primitive_set() -> None:
    """The in-package osa layer must expose every primitive the
    Word oracle (Spec 011 + later) consumes — same shape as
    `pptx.osa` and `xlsx.osa`."""
    assert _REQUIRED_OSA_SYMBOLS.issubset(set(docx_osa.__all__))


def test_docx_osa_symbols_resolve_to_callables_or_data() -> None:
    """Each exported name resolves on the module (no NameError)."""
    for name in _REQUIRED_OSA_SYMBOLS:
        assert hasattr(docx_osa, name), f"docx.osa missing {name}"


def test_docx_osa_has_callable_dialog_helpers() -> None:
    """The dialog-handling helpers must be callable functions."""
    for name in (
        "find_repair_dialog_text",
        "click_dialog_button",
        "dismiss_repair_dialog",
        "dismiss_any_leftover_modal",
        "is_word_running",
        "word_version",
        "open_document",
        "list_open_document_names",
        "is_document_open",
        "activate_document",
    ):
        assert callable(getattr(docx_osa, name)), f"{name} not callable"


def test_docx_osa_repair_dialog_patterns_documented() -> None:
    """The pattern list must be non-empty and include the 0.6.6
    additions (`unable to read this document` etc.) that closed the
    Word "hard-error" dialog gap."""
    patterns = docx_osa.REPAIR_DIALOG_PATTERNS
    assert len(patterns) >= 8
    assert "unreadable content" in patterns
    assert "unable to read this document" in patterns


def test_docx_osa_accept_button_labels_include_open_and_repair() -> None:
    """The accept-button list must include `Open and Repair` so the
    auto-dismiss path catches the hard-error dialog (added in 0.6.6
    after the baseline run surfaced the gap)."""
    labels = docx_osa.REPAIR_DIALOG_ACCEPT_BUTTON_LABELS
    assert len(labels) >= 4
    assert "Yes" in labels
    assert "Open and Repair" in labels


def test_docx_osa_close_document_saving_distinct_from_close_document() -> None:
    """The two close primitives are functionally distinct: `close_document`
    saves nothing, `close_document_saving` writes back to the open path
    via the close handler. The 0.7.4 consolidation must preserve both;
    if a future refactor accidentally collapses them, the matrix-driven
    Word oracle's persist path breaks."""
    assert docx_osa.close_document is not docx_osa.close_document_saving


def test_docx_osa_close_active_document_aliases_match() -> None:
    """Spec 011 historical names must alias to the canonical ones."""
    assert docx_osa.close_active_document is docx_osa.close_document
    assert docx_osa.close_active_document_saving is docx_osa.close_document_saving
    assert docx_osa.launch_word_app is docx_osa.launch_word


def test_dismiss_repair_dialog_signature_when_no_dialog() -> None:
    """`dismiss_repair_dialog()` returns `(False, None, None)` when no
    dialog is visible. Same contract as `pptx.osa.dismiss_repair_dialog`
    / `xlsx.osa.dismiss_repair_dialog`."""
    result = docx_osa.dismiss_repair_dialog()
    assert isinstance(result, tuple)
    assert len(result) == 3
    seen, text, clicked = result
    assert isinstance(seen, bool)
    assert text is None or isinstance(text, str)
    assert clicked is None or isinstance(clicked, str)


# -- back-compat shim contract -----------------------------------------------


def test_word_window_shim_reexports_all_symbols() -> None:
    """Existing imports from `tools.oracle.word_window` keep working —
    every symbol Spec 011 originally exposed should still resolve.
    If this breaks, all existing word oracle code breaks."""
    for name in _REQUIRED_OSA_SYMBOLS:
        assert hasattr(word_window, name), f"word_window shim missing {name}"


def test_word_window_shim_osascript_error_alias_catches_runtime_error() -> None:
    """Historical: `word_window.OsascriptError` was a RuntimeError
    subclass. In the new layer `openxml_audit.osa.osascript` raises
    plain RuntimeError; the shim aliases `OsascriptError = RuntimeError`
    so existing `except word_window.OsascriptError` clauses still
    catch what they caught. Verify the alias holds."""
    assert word_window.OsascriptError is RuntimeError
    # And it actually catches a raised RuntimeError:
    try:
        raise RuntimeError("dummy")
    except word_window.OsascriptError as exc:
        assert str(exc) == "dummy"


def test_word_window_shim_exposes_internal_applescript_quote() -> None:
    """`tools/oracle/word_roundtrip.py` and other consumers used
    `word_window._applescript_quote()` historically. The shim
    aliases it from `openxml_audit.osa.applescript_quote`."""
    assert hasattr(word_window, "_applescript_quote")
    assert word_window._applescript_quote is word_window.applescript_quote
