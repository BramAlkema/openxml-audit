"""Small, dependency-free client for the Euro-Office Document Server API.

The request shape mirrors Nextcloud connector 11.0.1.  JWT credentials are
read from the environment by :meth:`EuroOfficeClient.from_env`; the oracle
does not accept secrets on its command line, where process listings and shell
history could expose them.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Any, cast


class EuroOfficeError(Exception):
    """Base class for Euro-Office client failures."""


class EuroOfficeConfigError(EuroOfficeError):
    """Missing or invalid client configuration."""


class EuroOfficeRequestError(EuroOfficeError):
    """Transport or malformed-response failure."""


class EuroOfficeConversionError(EuroOfficeError):
    """Document Server rejected or failed a conversion."""


@dataclass(frozen=True)
class ConversionResult:
    """Successful synchronous conversion response."""

    file_url: str
    percent: int | None = None
    end_convert: bool | None = None


_CONVERSION_ERRORS = {
    -1: "unknown conversion error",
    -2: "conversion timed out",
    -3: "conversion failed",
    -4: "source document could not be downloaded",
    -5: "incorrect document password",
    -6: "conversion database error",
    -7: "invalid conversion input",
    -8: "invalid JWT token",
}

_USER_AGENT = "openxml-audit-eurooffice-oracle/1"


def _base64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _encode_jwt(payload: dict[str, Any], secret: str) -> str:
    """Encode an HS256 JWT without adding a third-party dependency."""

    header = {"alg": "HS256", "typ": "JWT"}
    header_part = _base64url(json.dumps(header, separators=(",", ":")).encode())
    payload_part = _base64url(json.dumps(payload, separators=(",", ":")).encode())
    signing_input = f"{header_part}.{payload_part}".encode("ascii")
    signature = hmac.new(secret.encode(), signing_input, hashlib.sha256).digest()
    return f"{header_part}.{payload_part}.{_base64url(signature)}"


def _endpoint_label(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    return parsed.path or "/"


class EuroOfficeClient:
    """Client for health, version, conversion, and artifact download calls."""

    def __init__(
        self,
        server_url: str,
        *,
        jwt_secret: str | None = None,
        jwt_header: str = "Authorization",
        timeout: float = 120.0,
    ) -> None:
        if not server_url.strip():
            raise EuroOfficeConfigError("Euro-Office server URL is required")
        parsed = urllib.parse.urlsplit(server_url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise EuroOfficeConfigError("Euro-Office server URL must use HTTP(S)")
        if not jwt_header.strip():
            raise EuroOfficeConfigError("JWT header name cannot be empty")
        if timeout <= 0:
            raise EuroOfficeConfigError("timeout must be greater than zero")

        self._server_url = server_url.rstrip("/") + "/"
        self._jwt_secret = jwt_secret
        self._jwt_header = jwt_header
        self._timeout = timeout

    @property
    def server_url(self) -> str:
        """Normalized, non-secret server URL."""

        return self._server_url

    @classmethod
    def from_env(
        cls,
        server_url: str | None = None,
        *,
        timeout: float = 120.0,
    ) -> EuroOfficeClient:
        """Build a client from explicit URL plus secret-safe environment vars."""

        resolved_url = server_url or os.environ.get("EUROOFFICE_ORACLE_URL")
        if not resolved_url:
            raise EuroOfficeConfigError("set EUROOFFICE_ORACLE_URL or pass --server-url")
        return cls(
            resolved_url,
            jwt_secret=os.environ.get("EUROOFFICE_ORACLE_JWT_SECRET"),
            jwt_header=os.environ.get("EUROOFFICE_ORACLE_JWT_HEADER", "Authorization"),
            timeout=timeout,
        )

    def _url(self, relative: str) -> str:
        return urllib.parse.urljoin(self._server_url, relative)

    def _request(
        self,
        url: str,
        *,
        body: bytes | None = None,
        headers: dict[str, str] | None = None,
        decode_json: bool = False,
    ) -> Any:
        method = "POST" if body is not None else "GET"
        request_headers = {"User-Agent": _USER_AGENT}
        request_headers.update(headers or {})
        request = urllib.request.Request(url, data=body, method=method, headers=request_headers)
        try:
            with urllib.request.urlopen(request, timeout=self._timeout) as response:
                payload = response.read()
        except urllib.error.HTTPError as exc:
            # Do not echo response bodies or signed query strings: either can contain
            # credentials supplied by the server or upstream source URL.
            raise EuroOfficeRequestError(
                f"Euro-Office {_endpoint_label(url)} returned HTTP {exc.code}"
            ) from exc
        except urllib.error.URLError as exc:
            raise EuroOfficeRequestError(
                f"Euro-Office {_endpoint_label(url)} request failed: {exc.reason}"
            ) from exc

        if not decode_json:
            return payload
        try:
            decoded = json.loads(payload)
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise EuroOfficeRequestError(
                f"Euro-Office {_endpoint_label(url)} returned invalid JSON"
            ) from exc
        if not isinstance(decoded, dict):
            raise EuroOfficeRequestError(
                f"Euro-Office {_endpoint_label(url)} returned a non-object response"
            )
        return decoded

    def _authenticated_payload(
        self,
        data: dict[str, Any],
    ) -> tuple[dict[str, Any], dict[str, str]]:
        headers = {"Content-Type": "application/json"}
        if not self._jwt_secret:
            return data, headers

        now = int(time.time())
        expires = now + 300
        body_data = {**data, "iat": now, "exp": expires}
        body_data["token"] = _encode_jwt(body_data, self._jwt_secret)
        outer_token = _encode_jwt(
            {"payload": body_data, "iat": now, "exp": expires},
            self._jwt_secret,
        )
        prefix = "Bearer " if self._jwt_header.lower() == "authorization" else ""
        headers[self._jwt_header] = f"{prefix}{outer_token}"
        return body_data, headers

    def healthcheck(self) -> bool:
        """Return whether ``/healthcheck`` reports a healthy server."""

        payload = cast(bytes, self._request(self._url("healthcheck")))
        return payload.decode("utf-8", errors="replace").strip().lower() == "true"

    def version(self) -> str:
        """Return the Document Server version from CommandService."""

        data, headers = self._authenticated_payload({"c": "version"})
        payload = self._request(
            self._url("coauthoring/CommandService.ashx"),
            body=json.dumps(data, separators=(",", ":")).encode(),
            headers=headers,
            decode_json=True,
        )
        error = payload.get("error")
        if isinstance(error, int) and error != 0:
            raise EuroOfficeRequestError(f"Euro-Office version request returned error {error}")
        version = payload.get("version")
        if not isinstance(version, str) or not version:
            raise EuroOfficeRequestError("Euro-Office version response omitted version")
        build = payload.get("buildNumber")
        return f"{version}.{build}" if isinstance(build, int | str) and str(build) else version

    def convert(
        self,
        *,
        source_url: str,
        source_format: str,
        target_format: str,
        key: str,
        title: str | None = None,
    ) -> ConversionResult:
        """Synchronously convert a URL-addressable document."""

        if not 1 <= len(key) <= 20 or not all(char.isalnum() or char in "-._=" for char in key):
            raise EuroOfficeConfigError(
                "conversion key must be 1-20 ASCII letters, digits, or -._="
            )
        source_format = source_format.lower().lstrip(".")
        target_format = target_format.lower().lstrip(".")
        data: dict[str, Any] = {
            "async": False,
            "url": source_url,
            "outputtype": target_format,
            "filetype": source_format,
            "title": title or f"document.{source_format}",
            "key": key,
        }
        request_data, headers = self._authenticated_payload(data)
        converter_url = self._url("converter?" + urllib.parse.urlencode({"shardKey": key}))
        payload = self._request(
            converter_url,
            body=json.dumps(request_data, separators=(",", ":")).encode(),
            headers=headers,
            decode_json=True,
        )

        error = payload.get("error", 0)
        if isinstance(error, int) and error != 0:
            message = _CONVERSION_ERRORS.get(error, "unrecognized conversion error")
            raise EuroOfficeConversionError(f"Euro-Office conversion error {error}: {message}")
        file_url = payload.get("fileUrl")
        if not isinstance(file_url, str) or not file_url:
            raise EuroOfficeRequestError("Euro-Office conversion response omitted fileUrl")
        percent = payload.get("percent")
        end_convert = payload.get("endConvert")
        return ConversionResult(
            file_url=file_url,
            percent=percent if isinstance(percent, int) else None,
            end_convert=end_convert if isinstance(end_convert, bool) else None,
        )

    def download(self, file_url: str) -> bytes:
        """Download a conversion artifact returned by Document Server."""

        return cast(bytes, self._request(file_url))


__all__ = [
    "ConversionResult",
    "EuroOfficeClient",
    "EuroOfficeConfigError",
    "EuroOfficeConversionError",
    "EuroOfficeError",
    "EuroOfficeRequestError",
]
