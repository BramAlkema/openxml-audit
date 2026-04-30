"""Back-compat shim — Word window primitives now live in
`openxml_audit.docx.osa`.

This module pre-dates the in-package osa consolidation (Spec 030,
0.7.4). Spec 011 (Word roundtrip oracle, 0.5.0) shipped the
primitives here directly; subsequent format additions
(`pptx.osa` for PPTX in 0.6.4, `xlsx.osa` for XLSX in 0.6.5) put
their primitives in the in-package layer for symmetry. 0.7.4
finished that consolidation: every Word primitive lives in
`openxml_audit.docx.osa` now.

Existing consumers (`tools/oracle/word_roundtrip.py`,
`tools/oracle/word_repair_oracle.py`,
`tools/oracle/word_repair_corpus.py`) keep working through this
shim. Future code should import directly from
`openxml_audit.docx.osa`.
"""

from __future__ import annotations

# Re-export everything the old word_window.py exposed. The new
# canonical home is openxml_audit.docx.osa; aliases below preserve
# the old historical names where they differed.
from openxml_audit.docx.osa import (  # noqa: F401  (re-exports)
    REPAIR_DIALOG_ACCEPT_BUTTON_LABELS,
    REPAIR_DIALOG_PATTERNS,
    WORD_APP_BUNDLE,
    WORD_APP_ID,
    WORD_PROCESS_NAME,
    activate_document,
    click_dialog_button,
    close_active_document,
    close_active_document_saving,
    close_document,
    close_document_saving,
    dismiss_any_leftover_modal,
    dismiss_repair_dialog,
    find_repair_dialog_text,
    is_document_open,
    is_word_running,
    launch_word,
    launch_word_app,
    list_open_document_names,
    open_document,
    save_document,
    word_version,
)
from openxml_audit.osa import (  # noqa: F401
    applescript_quote,
    osascript,
    osascript_jxa,
)


# Internal alias used by historical callers.
_applescript_quote = applescript_quote


# Historical: `word_window.OsascriptError` was a `RuntimeError`
# subclass defined here. In the new layer,
# `openxml_audit.osa.osascript` raises plain `RuntimeError`; aliasing
# `OsascriptError = RuntimeError` here means existing
# `except word_window.OsascriptError` clauses still catch what they
# always caught.
OsascriptError = RuntimeError


__all__ = [
    "REPAIR_DIALOG_ACCEPT_BUTTON_LABELS",
    "REPAIR_DIALOG_PATTERNS",
    "WORD_APP_BUNDLE",
    "WORD_APP_ID",
    "WORD_PROCESS_NAME",
    "OsascriptError",
    "activate_document",
    "click_dialog_button",
    "close_active_document",
    "close_active_document_saving",
    "close_document",
    "close_document_saving",
    "dismiss_any_leftover_modal",
    "dismiss_repair_dialog",
    "find_repair_dialog_text",
    "is_document_open",
    "is_word_running",
    "launch_word",
    "launch_word_app",
    "list_open_document_names",
    "open_document",
    "save_document",
    "word_version",
    "osascript",
    "osascript_jxa",
    "applescript_quote",
]
