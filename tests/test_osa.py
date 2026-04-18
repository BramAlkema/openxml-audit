"""Pure-Python tests for osa primitives. The subprocess-backed functions
are not unit-tested — they require macOS + the target app to exercise.
"""

from openxml_audit.osa import applescript_quote


def test_applescript_quote_wraps_plain_string() -> None:
    assert applescript_quote("hello world") == '"hello world"'


def test_applescript_quote_escapes_double_quotes() -> None:
    assert applescript_quote('say "hi"') == '"say \\"hi\\""'


def test_applescript_quote_escapes_backslashes() -> None:
    assert applescript_quote(r"C:\Users") == '"C:\\\\Users"'


def test_applescript_quote_handles_empty_string() -> None:
    assert applescript_quote("") == '""'
