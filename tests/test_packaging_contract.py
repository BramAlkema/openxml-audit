"""Regression checks for files required by installed console scripts."""

from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]


def test_wheel_includes_oracle_dispatcher_implementations() -> None:
    """The public oracle entry point imports a package outside ``src``.

    Keep the Hatch force-include mapping explicit: source-checkout tests can
    import ``tools.oracle`` from the repository root and would otherwise miss
    the broken-wheel regression.
    """
    pyproject = (REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8")

    assert '"tools/oracle" = "tools/oracle"' in pyproject
