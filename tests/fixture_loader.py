"""Shared fixture loading helpers for tests."""

from __future__ import annotations

from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def load_fixture_bytes(*parts: str) -> bytes:
    """Load fixture contents as raw bytes."""
    return (FIXTURES_DIR.joinpath(*parts)).read_bytes()


def load_fixture_text(*parts: str) -> str:
    """Load fixture contents as text."""
    return load_fixture_bytes(*parts).decode("utf-8")
