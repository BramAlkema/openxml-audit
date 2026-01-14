"""Tests for binary payload validation."""

from __future__ import annotations

from openxml_audit.binary import parse_font_key, validate_binary_content
from openxml_audit.errors import ValidationSeverity
from tests.fixture_loader import load_fixture_bytes


def test_valid_png_signature() -> None:
    data = load_fixture_bytes("binary", "valid.png")
    assert validate_binary_content("image/png", "/ppt/media/image1.png", data) is None


def test_invalid_png_signature() -> None:
    data = load_fixture_bytes("binary", "invalid.png")
    assert validate_binary_content("image/png", "/ppt/media/image1.png", data) is not None


def test_valid_jpeg_signature() -> None:
    data = load_fixture_bytes("binary", "valid.jpg")
    assert validate_binary_content("image/jpeg", "/ppt/media/image1.jpg", data) is None


def test_invalid_jpeg_signature() -> None:
    data = load_fixture_bytes("binary", "invalid.jpg")
    assert validate_binary_content("image/jpeg", "/ppt/media/image1.jpg", data) is not None


def test_valid_emf_signature() -> None:
    data = load_fixture_bytes("binary", "valid.emf")
    assert validate_binary_content("image/emf", "/ppt/media/image1.emf", data) is None


def test_valid_wmf_signature() -> None:
    data = load_fixture_bytes("binary", "valid.wmf")
    assert validate_binary_content("image/wmf", "/ppt/media/image1.wmf", data) is None


def test_valid_ole_signature() -> None:
    data = load_fixture_bytes("binary", "valid.ole")
    assert (
        validate_binary_content(
            "application/vnd.openxmlformats-officedocument.oleObject",
            "/word/embeddings/oleObject1.bin",
            data,
        )
        is None
    )


def test_invalid_ole_signature() -> None:
    data = load_fixture_bytes("binary", "invalid.ole")
    assert validate_binary_content(
        "application/vnd.openxmlformats-officedocument.oleObject",
        "/word/embeddings/oleObject1.bin",
        data,
    ) is not None


def test_valid_ttf_signature() -> None:
    data = load_fixture_bytes("binary", "valid.ttf")
    assert (
        validate_binary_content(
            "application/x-font-ttf",
            "/ppt/embeddings/font1.ttf",
            data,
        )
        is None
    )


def test_invalid_ttf_signature() -> None:
    data = load_fixture_bytes("binary", "invalid.ttf")
    assert (
        validate_binary_content(
            "application/x-font-ttf",
            "/ppt/embeddings/font1.ttf",
            data,
        )
        is not None
    )


def test_valid_obfuscated_font_signature() -> None:
    data = load_fixture_bytes("binary", "valid.odttf")
    key = parse_font_key("{00112233-4455-6677-8899-AABBCCDDEEFF}")
    assert key is not None
    assert (
        validate_binary_content(
            "application/vnd.openxmlformats-officedocument.obfuscatedFont",
            "/word/fonts/font1.odttf",
            data,
            font_key=key,
        )
        is None
    )


def test_obfuscated_font_missing_key_warns() -> None:
    data = load_fixture_bytes("binary", "valid.odttf")
    result = validate_binary_content(
        "application/vnd.openxmlformats-officedocument.obfuscatedFont",
        "/word/fonts/font1.odttf",
        data,
    )
    assert result is not None
    assert result.severity == ValidationSeverity.WARNING
