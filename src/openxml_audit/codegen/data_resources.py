"""Resolve bundled Open XML SDK data for both wheels and source checkouts."""

from __future__ import annotations

from importlib.resources import files
from os import PathLike, fspath
from pathlib import Path

PROJECT_DATA_DIR = Path(__file__).resolve().parents[3] / "data" / "openxml"


def get_openxml_data_dir() -> Path:
    """Return the Open XML SDK data directory for the current environment."""
    packaged_data_dir = files("openxml_audit").joinpath("data").joinpath("openxml")
    if isinstance(packaged_data_dir, PathLike):
        packaged_path = Path(fspath(packaged_data_dir))
        if packaged_path.is_dir():
            return packaged_path

    return PROJECT_DATA_DIR
