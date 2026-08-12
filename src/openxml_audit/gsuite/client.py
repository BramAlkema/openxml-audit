"""Drive API client for the GSuite roundtrip oracle.

Spec 031. Four primitives — upload, convert-to-native, export-to-ooxml,
delete — wired to the Drive REST API v3 with stdlib `urllib.request`.

Auth: `google-auth` handles the JWT signing + token refresh dance for
service accounts (RSA-SHA256 over PKCS#8 keys, claims, exchange at
`oauth2.googleapis.com/token`). Everything else is plain HTTP, so we
don't depend on `google-api-python-client` or its 3 MB transitive tree
(`requests`, `urllib3`, `protobuf`, `httplib2`, `uritemplate`,
`googleapis-common-protos`, ...). One minor wrinkle: `google-auth`
expects a "transport" callable for `refresh()`. We satisfy that with
`_UrllibAuthTransport` below — ~30 lines that wraps `urllib.request`
in the shape google-auth wants.

Service accounts have zero storage quota since Google's 2024 policy
change, so `with_subject(...)` impersonation of a real Workspace user
via domain-wide delegation is mandatory for any `files.create` call.
See `specs/031-gsuite-roundtrip-oracle.md` for the one-time admin-
console setup.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from google.oauth2.service_account import Credentials

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

SLIDES_MIME = "application/vnd.google-apps.presentation"
DOC_MIME = "application/vnd.google-apps.document"
SHEET_MIME = "application/vnd.google-apps.spreadsheet"

DEFAULT_CREDS_PATH = Path.home() / ".config" / "openxml-audit" / "google_service_account.json"
DEFAULT_SCOPES = ("https://www.googleapis.com/auth/drive",)

_DRIVE_API = "https://www.googleapis.com/drive/v3"
_UPLOAD_API = "https://www.googleapis.com/upload/drive/v3"


class GSuiteError(Exception):
    """Base class for gsuite client errors."""


class GSuiteAuthError(GSuiteError):
    """Authentication or authorization failure (missing creds, no
    delegation, missing dep, etc.)."""


class _UrllibAuthResponse:
    """Shape google-auth's `Response` interface expects."""

    def __init__(self, status: int, headers: Any, data: bytes) -> None:
        self.status = status
        self.headers = headers
        self.data = data


class _UrllibAuthTransport:
    """Minimal transport for google-auth's `refresh()`.

    google-auth's service-account flow calls `request(url, method, body,
    headers, timeout)` to exchange a signed JWT for an access token.
    The full transport interface is small; we implement just `__call__`
    using `urllib.request` so we don't pull in `requests` or `urllib3`.
    Only invoked during token refresh (~hourly), so performance is a
    non-issue.
    """

    def __call__(
        self,
        url: str,
        method: str = "GET",
        body: bytes | None = None,
        headers: dict[str, str] | None = None,
        timeout: float | None = 60.0,
        **_kwargs: Any,
    ) -> _UrllibAuthResponse:
        req = urllib.request.Request(url, data=body, method=method)
        for k, v in (headers or {}).items():
            req.add_header(k, v)
        try:
            resp = urllib.request.urlopen(req, timeout=timeout)
            return _UrllibAuthResponse(resp.status, resp.headers, resp.read())
        except urllib.error.HTTPError as exc:
            # google-auth inspects .status/.data to decide whether to
            # retry or raise. Surface the failure shape it expects.
            return _UrllibAuthResponse(exc.code, exc.headers, exc.read())


class GSuiteClient:
    """Thin Drive API wrapper for the roundtrip oracle.

    Construct via `GSuiteClient.from_service_account(...)`; the raw
    constructor is for tests that pass an in-memory credentials object.
    """

    def __init__(
        self,
        credentials: Credentials,
        *,
        subject: str | None = None,
    ) -> None:
        self._creds = credentials
        self._subject = subject
        self._auth_transport = _UrllibAuthTransport()

    @property
    def subject(self) -> str | None:
        """The impersonated Workspace user, if delegation is in use."""
        return self._subject

    @property
    def credentials(self) -> Credentials:
        """The underlying google-auth `Credentials` instance."""
        return self._creds

    @classmethod
    def from_service_account(
        cls,
        creds_path: Path | str | None = None,
        *,
        subject: str | None = None,
        scopes: tuple[str, ...] = DEFAULT_SCOPES,
    ) -> GSuiteClient:
        """Build a client from a service-account JSON key.

        `subject` is the Workspace user the SA should impersonate via
        domain-wide delegation. If unset, falls back to the
        `GSUITE_ORACLE_SUBJECT` env var. Without a subject the SA has
        no storage quota and `upload(...)` will fail with
        `storageQuotaExceeded`.

        `creds_path` falls back to `GSUITE_ORACLE_CREDS` env var, then
        to `~/.config/openxml-audit/google_service_account.json`.
        """
        try:
            from google.oauth2 import service_account
        except ImportError as exc:
            raise GSuiteAuthError(
                'google-auth not installed. Run `pip install -e ".[gsuite]"` to install it.'
            ) from exc

        path = Path(
            creds_path or os.environ.get("GSUITE_ORACLE_CREDS") or DEFAULT_CREDS_PATH
        ).expanduser()
        if not path.exists():
            raise GSuiteAuthError(
                f"Service account JSON not found at {path}. See "
                f"specs/031-gsuite-roundtrip-oracle.md for setup."
            )

        resolved_subject = subject or os.environ.get("GSUITE_ORACLE_SUBJECT")

        try:
            base = service_account.Credentials.from_service_account_file(
                str(path), scopes=list(scopes)
            )
            creds = base.with_subject(resolved_subject) if resolved_subject else base
        except Exception as exc:
            raise GSuiteAuthError(f"Drive auth setup failed: {exc}") from exc

        return cls(creds, subject=resolved_subject)

    # --- HTTP plumbing ------------------------------------------------------

    def _ensure_token(self) -> str:
        """Return a fresh access token, refreshing if needed."""
        if not self._creds.valid:
            try:
                self._creds.refresh(self._auth_transport)
            except Exception as exc:
                raise GSuiteAuthError(f"token refresh failed: {exc}") from exc
        return self._creds.token

    def _request(
        self,
        method: str,
        url: str,
        *,
        headers: dict[str, str] | None = None,
        body: bytes | None = None,
        decode_json: bool = True,
        timeout: float = 60.0,
    ) -> Any:
        """Make a Drive API request and return parsed JSON (or raw bytes
        when `decode_json=False`).

        Raises `GSuiteError` with the API's error payload preserved in
        the message so the oracle's notes/regression-tests can pattern-
        match on Google's wording (e.g. "Internal Error", "500").
        """
        token = self._ensure_token()
        full_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            full_headers.update(headers)
        req = urllib.request.Request(url, data=body, method=method, headers=full_headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                payload = resp.read()
        except urllib.error.HTTPError as exc:
            err_body = exc.read().decode("utf-8", errors="replace")
            raise GSuiteError(
                f"<HttpError {exc.code} when requesting {url} returned "
                f"{exc.reason!r}. Details: {err_body!r}>"
            ) from exc
        except urllib.error.URLError as exc:
            raise GSuiteError(f"network error for {method} {url}: {exc}") from exc

        if not decode_json:
            return payload
        if not payload:
            return None
        return json.loads(payload)

    # --- Public API ---------------------------------------------------------

    def whoami(self) -> dict[str, str]:
        """Return `{emailAddress, displayName}` of the active principal.

        Useful as a delegation smoke test: if domain-wide delegation
        isn't set up correctly this call surfaces the error before any
        upload is attempted.
        """
        result = self._request(
            "GET",
            f"{_DRIVE_API}/about?fields=" + urllib.parse.quote("user(emailAddress,displayName)"),
        )
        user = (result or {}).get("user", {})
        return {
            "emailAddress": user.get("emailAddress", ""),
            "displayName": user.get("displayName", ""),
        }

    def upload(
        self,
        path: Path | str,
        *,
        parent_id: str | None = None,
        mime_type: str | None = None,
        name: str | None = None,
    ) -> str:
        """Upload `path` to Drive; return the new file's ID.

        `parent_id` should be a folder owned by the impersonation
        subject (or a Shared Drive folder the SA can write to).
        `mime_type` defaults to the OOXML mime matching the file
        suffix; pass explicitly to override.

        Uses the Drive API's simple multipart upload (one HTTP POST,
        body = metadata JSON + file bytes). For files >5 MB consider
        switching to resumable uploads — out of scope for Phase 1
        oracle inputs (typically <100 KB pptx).
        """
        path = Path(path)
        if not path.exists():
            raise GSuiteError(f"upload source not found: {path}")

        resolved_mime = mime_type or _mime_for_suffix(path.suffix)
        metadata: dict[str, Any] = {"name": name or path.name}
        if parent_id:
            metadata["parents"] = [parent_id]

        body, content_type = _build_multipart(metadata, path, resolved_mime)
        result = self._request(
            "POST",
            f"{_UPLOAD_API}/files?uploadType=multipart&fields=id",
            headers={"Content-Type": content_type},
            body=body,
        )
        return result["id"]

    def convert_to_native(
        self,
        file_id: str,
        *,
        target_mime: str,
        parent_id: str | None = None,
        name: str | None = None,
    ) -> str:
        """Copy `file_id` with conversion to a native Google mime
        (`SLIDES_MIME`, `DOC_MIME`, or `SHEET_MIME`); return the new
        file's ID.

        This is the lossy import step — Google maps the OOXML into
        its proprietary IR, dropping or transforming features that
        don't fit.
        """
        body: dict[str, Any] = {"mimeType": target_mime}
        if name:
            body["name"] = name
        if parent_id:
            body["parents"] = [parent_id]
        result = self._request(
            "POST",
            f"{_DRIVE_API}/files/{file_id}/copy?fields=" + urllib.parse.quote("id,mimeType"),
            headers={"Content-Type": "application/json; charset=UTF-8"},
            body=json.dumps(body).encode("utf-8"),
        )
        return result["id"]

    def export_to_ooxml(self, file_id: str, ooxml_mime: str) -> bytes:
        """Export a native Google file back to OOXML bytes.

        This is the second lossy step — Google reconstructs OOXML
        from its IR, generally producing a larger, normalized
        package with added structure (notes masters, extra theme
        variants) and stripped metadata.
        """
        url = f"{_DRIVE_API}/files/{file_id}/export?mimeType={urllib.parse.quote(ooxml_mime)}"
        return self._request("GET", url, decode_json=False)

    def delete(self, file_id: str) -> bool:
        """Best-effort delete; returns True on success, False on
        failure (logs nothing — caller decides how to handle)."""
        try:
            self._request(
                "DELETE",
                f"{_DRIVE_API}/files/{file_id}",
                decode_json=False,
            )
            return True
        except GSuiteError:
            return False


# --- multipart upload body --------------------------------------------------


def _build_multipart(
    metadata: dict[str, Any],
    file_path: Path,
    mime_type: str,
) -> tuple[bytes, str]:
    """Build a `multipart/related` body for Drive's simple upload.

    The format Drive expects:

        --<boundary>
        Content-Type: application/json; charset=UTF-8

        {"name": "...", "parents": [...]}
        --<boundary>
        Content-Type: <mime_type>

        <raw file bytes>
        --<boundary>--

    Returns `(body, content_type)` ready to pass to `_request`.
    """
    boundary = "==boundary==" + uuid.uuid4().hex
    metadata_json = json.dumps(metadata).encode("utf-8")
    file_bytes = file_path.read_bytes()
    parts = [
        f"--{boundary}\r\n".encode(),
        b"Content-Type: application/json; charset=UTF-8\r\n\r\n",
        metadata_json,
        f"\r\n--{boundary}\r\n".encode(),
        f"Content-Type: {mime_type}\r\n\r\n".encode(),
        file_bytes,
        f"\r\n--{boundary}--\r\n".encode(),
    ]
    body = b"".join(parts)
    return body, f"multipart/related; boundary={boundary}"


# --- mime helpers -----------------------------------------------------------


def _mime_for_suffix(suffix: str) -> str:
    """Map a file extension to its OOXML mime type."""
    mapping = {
        ".pptx": PPTX_MIME,
        ".docx": DOCX_MIME,
        ".xlsx": XLSX_MIME,
    }
    mime = mapping.get(suffix.lower())
    if mime is None:
        raise GSuiteError(f"no default mime for suffix {suffix!r}; pass mime_type explicitly")
    return mime


def native_mime_for(target_format: str) -> str:
    """Map a target-format string (`pptx`/`docx`/`xlsx`) to the
    matching native Google mime type."""
    mapping = {
        "pptx": SLIDES_MIME,
        "docx": DOC_MIME,
        "xlsx": SHEET_MIME,
    }
    mime = mapping.get(target_format.lower())
    if mime is None:
        raise GSuiteError(f"unsupported target_format: {target_format!r}")
    return mime


def ooxml_mime_for(target_format: str) -> str:
    """Map a target-format string (`pptx`/`docx`/`xlsx`) to the
    matching OOXML mime type."""
    mapping = {
        "pptx": PPTX_MIME,
        "docx": DOCX_MIME,
        "xlsx": XLSX_MIME,
    }
    mime = mapping.get(target_format.lower())
    if mime is None:
        raise GSuiteError(f"unsupported target_format: {target_format!r}")
    return mime
