"""Euro-Office format matrix and Document Server client contract tests."""

from __future__ import annotations

import base64
import json
import urllib.error
import urllib.request
from typing import Any

import pytest

from openxml_audit.eurooffice import (
    EuroOfficeClient,
    EuroOfficeConversionError,
    EuroOfficeFormatMode,
    EuroOfficeRequestError,
    format_support,
)
from openxml_audit.eurooffice.client import _encode_jwt


@pytest.mark.parametrize(
    ("extension", "document_type"),
    [
        *((extension, "word") for extension in ("docx", "docm", "dotx", "dotm")),
        *((extension, "cell") for extension in ("xlsx", "xlsm", "xltx", "xltm")),
        *((extension, "slide") for extension in ("pptx", "pptm", "potx", "potm", "ppsx", "ppsm")),
    ],
)
def test_native_edit_matrix(extension: str, document_type: str) -> None:
    support = format_support(extension)
    assert support.document_type == document_type
    assert support.mode is EuroOfficeFormatMode.NATIVE_EDIT
    assert support.conversion_target == extension
    assert support.actions == ("view", "edit")


@pytest.mark.parametrize(
    ("extension", "document_type", "target"),
    [
        ("odt", "word", "docx"),
        ("ott", "word", "docx"),
        ("ods", "cell", "xlsx"),
        ("ots", "cell", "xlsx"),
        ("odp", "slide", "pptx"),
        ("otp", "slide", "pptx"),
    ],
)
def test_lossy_edit_matrix(extension: str, document_type: str, target: str) -> None:
    support = format_support(f"EXAMPLE.{extension.upper()}")
    assert support.document_type == document_type
    assert support.mode is EuroOfficeFormatMode.LOSSY_EDIT
    assert support.conversion_target == target
    assert support.actions == ("view", "edit", "auto-convert")


def test_odg_is_view_convert_only() -> None:
    support = format_support("https://files.example.test/drawing.ODG?download=1")
    assert support.mode is EuroOfficeFormatMode.VIEW_ONLY
    assert support.conversion_target == "pptx"
    assert support.actions == ("view", "auto-convert")


@pytest.mark.parametrize("extension", ["odm", "xlam", "thmx", "ppam", "txt", ""])
def test_unsupported_formats_are_explicit(extension: str) -> None:
    support = format_support(extension)
    assert support.mode is EuroOfficeFormatMode.UNSUPPORTED
    assert support.conversion_target is None
    assert support.actions == ()


def _decode_jwt_part(token: str, part: int) -> dict[str, Any]:
    encoded = token.split(".")[part]
    padding = "=" * (-len(encoded) % 4)
    return json.loads(base64.urlsafe_b64decode(encoded + padding))


class _Response:
    def __init__(self, payload: bytes) -> None:
        self._payload = payload

    def __enter__(self) -> _Response:
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def read(self) -> bytes:
        return self._payload


def test_convert_matches_connector_jwt_request_shape(monkeypatch: pytest.MonkeyPatch) -> None:
    captured: dict[str, Any] = {}

    def fake_urlopen(request: urllib.request.Request, *, timeout: float) -> _Response:
        captured["request"] = request
        captured["timeout"] = timeout
        return _Response(
            json.dumps({"error": 0, "fileUrl": "https://result.test/out.docx"}).encode()
        )

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setattr("openxml_audit.eurooffice.client.time.time", lambda: 1_700_000_000)
    client = EuroOfficeClient(
        "https://office.example.test/",
        jwt_secret="test-secret",
        timeout=42,
    )
    result = client.convert(
        source_url="https://files.example.test/input.odt",
        source_format="odt",
        target_format="docx",
        key="0123456789abcdef0123",
        title="input.odt",
    )

    request = captured["request"]
    assert isinstance(request, urllib.request.Request)
    assert request.full_url == (
        "https://office.example.test/converter?shardKey=0123456789abcdef0123"
    )
    assert captured["timeout"] == 42
    assert request.get_method() == "POST"
    assert request.headers["User-agent"] == "openxml-audit-eurooffice-oracle/1"
    assert request.headers["Content-type"] == "application/json"
    assert request.headers["Authorization"].startswith("Bearer ")

    body = json.loads(request.data)
    assert body["async"] is False
    assert body["filetype"] == "odt"
    assert body["outputtype"] == "docx"
    assert body["iat"] == 1_700_000_000
    assert body["exp"] == 1_700_000_300
    body_claims = {key: value for key, value in body.items() if key != "token"}
    assert body["token"] == _encode_jwt(body_claims, "test-secret")

    outer = request.headers["Authorization"].removeprefix("Bearer ")
    outer_claims = _decode_jwt_part(outer, 1)
    assert outer_claims == {
        "payload": body,
        "iat": 1_700_000_000,
        "exp": 1_700_000_300,
    }
    assert result.file_url == "https://result.test/out.docx"


def test_healthcheck_and_version(monkeypatch: pytest.MonkeyPatch) -> None:
    responses = iter(
        [
            _Response(b"true"),
            _Response(json.dumps({"error": 0, "version": "9.3.3", "buildNumber": 8}).encode()),
        ]
    )
    monkeypatch.setattr(urllib.request, "urlopen", lambda *_args, **_kwargs: next(responses))
    client = EuroOfficeClient("https://office.example.test")
    assert client.healthcheck()
    assert client.version() == "9.3.3.8"


def test_conversion_error_is_classified(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        urllib.request,
        "urlopen",
        lambda *_args, **_kwargs: _Response(b'{"error":-4}'),
    )
    client = EuroOfficeClient("https://office.example.test")
    with pytest.raises(EuroOfficeConversionError, match="could not be downloaded"):
        client.convert(
            source_url="https://files.example.test/missing.odt",
            source_format="odt",
            target_format="docx",
            key="valid-key",
        )


def test_http_error_does_not_leak_query_or_body(monkeypatch: pytest.MonkeyPatch) -> None:
    secret = "do-not-leak-this"

    def fail(*_args: object, **_kwargs: object) -> _Response:
        raise urllib.error.HTTPError(
            f"https://result.test/file?token={secret}",
            403,
            "Forbidden",
            {},
            None,
        )

    monkeypatch.setattr(urllib.request, "urlopen", fail)
    client = EuroOfficeClient("https://office.example.test", jwt_secret=secret)
    with pytest.raises(EuroOfficeRequestError) as caught:
        client.download(f"https://result.test/file?token={secret}")
    assert secret not in str(caught.value)
    assert "?" not in str(caught.value)
