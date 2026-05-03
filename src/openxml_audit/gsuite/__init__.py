"""Google Workspace integration for the roundtrip oracle.

Spec 031. Provides a thin Drive API wrapper used by
`tools/oracle/gsuite_roundtrip.py` to upload OOXML files, convert
them to native Google formats (Slides/Docs/Sheets), export back to
OOXML, and clean up — the four primitives an oracle needs.

Auth model: service account with domain-wide delegation. See
`specs/031-gsuite-roundtrip-oracle.md` for the one-time setup steps.

Optional dependency: `pip install -e ".[gsuite]"`.
"""

from openxml_audit.gsuite.client import (
    DEFAULT_CREDS_PATH,
    DEFAULT_SCOPES,
    DOC_MIME,
    DOCX_MIME,
    PPTX_MIME,
    SHEET_MIME,
    SLIDES_MIME,
    XLSX_MIME,
    GSuiteAuthError,
    GSuiteClient,
    GSuiteError,
)

__all__ = [
    "GSuiteClient",
    "GSuiteAuthError",
    "GSuiteError",
    "PPTX_MIME",
    "DOCX_MIME",
    "XLSX_MIME",
    "SLIDES_MIME",
    "DOC_MIME",
    "SHEET_MIME",
    "DEFAULT_CREDS_PATH",
    "DEFAULT_SCOPES",
]
