#!/usr/bin/env python3
"""Sync Open XML SDK data files from Microsoft's GitHub repository.

Downloads the schema definitions and schematron rules needed to generate
validation constraints.
"""

from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import urllib.request
from pathlib import Path
from typing import TYPE_CHECKING

# Configuration
SDK_REPO = "dotnet/Open-XML-SDK"
SDK_BRANCH = "main"
SDK_BASE_URL = f"https://raw.githubusercontent.com/{SDK_REPO}/{SDK_BRANCH}"
SDK_API_URL = f"https://api.github.com/repos/{SDK_REPO}"

# Files to download
DATA_FILES = [
    "data/namespaces.json",
    "data/schematrons.json",
]

# Project paths
PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "data" / "openxml"
SCHEMAS_DIR = DATA_DIR / "schemas"


def get_latest_commit() -> str:
    """Get the latest commit hash from the SDK repository."""
    url = f"{SDK_API_URL}/commits/{SDK_BRANCH}"
    req = urllib.request.Request(url)
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("User-Agent", "openxml-audit")

    with urllib.request.urlopen(req, timeout=30) as response:
        data = json.loads(response.read().decode())
        return data["sha"]


def get_current_version() -> str | None:
    """Get the currently synced SDK version."""
    version_file = DATA_DIR / ".sdk_version"
    if version_file.exists():
        return version_file.read_text().strip()
    return None


def save_version(commit_hash: str) -> None:
    """Save the synced SDK version."""
    version_file = DATA_DIR / ".sdk_version"
    version_file.write_text(commit_hash)


def download_file(url: str, dest: Path) -> None:
    """Download a file from URL to destination."""
    req = urllib.request.Request(url)
    req.add_header("User-Agent", "openxml-audit")

    with urllib.request.urlopen(req, timeout=60) as response:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(response.read())


def get_schema_file_list() -> list[str]:
    """Get list of schema files from the SDK repository."""
    url = f"{SDK_API_URL}/contents/data/schemas"
    req = urllib.request.Request(url)
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("User-Agent", "openxml-audit")

    with urllib.request.urlopen(req, timeout=30) as response:
        data = json.loads(response.read().decode())
        return [item["name"] for item in data if item["name"].endswith(".json")]


def sync_data(force: bool = False) -> dict[str, int]:
    """Sync data files from Open XML SDK.

    Args:
        force: If True, re-download even if already up to date.

    Returns:
        Dictionary with counts of downloaded files.
    """
    print("Checking Open XML SDK repository...")

    latest_commit = get_latest_commit()
    current_version = get_current_version()

    print(f"  Latest commit: {latest_commit[:12]}")
    print(f"  Current version: {current_version[:12] if current_version else 'none'}")

    if current_version == latest_commit and not force:
        print("Already up to date!")
        return {"schemas": 0, "data_files": 0, "skipped": True}

    # Create data directory
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SCHEMAS_DIR.mkdir(parents=True, exist_ok=True)

    stats = {"schemas": 0, "data_files": 0, "skipped": False}

    # Download data files
    print("\nDownloading data files...")
    for file_path in DATA_FILES:
        url = f"{SDK_BASE_URL}/{file_path}"
        dest = DATA_DIR / Path(file_path).name
        print(f"  {file_path}...", end=" ", flush=True)
        try:
            download_file(url, dest)
            print("OK")
            stats["data_files"] += 1
        except Exception as e:
            print(f"FAILED: {e}")

    # Download schema files
    print("\nDownloading schema files...")
    schema_files = get_schema_file_list()
    print(f"  Found {len(schema_files)} schema files")

    for i, filename in enumerate(schema_files, 1):
        url = f"{SDK_BASE_URL}/data/schemas/{filename}"
        dest = SCHEMAS_DIR / filename
        print(f"  [{i}/{len(schema_files)}] {filename}...", end=" ", flush=True)
        try:
            download_file(url, dest)
            print("OK")
            stats["schemas"] += 1
        except Exception as e:
            print(f"FAILED: {e}")

    # Save version
    save_version(latest_commit)

    print(f"\nSync complete!")
    print(f"  Schema files: {stats['schemas']}")
    print(f"  Data files: {stats['data_files']}")
    print(f"  SDK version: {latest_commit[:12]}")

    return stats


def main() -> None:
    """Main entry point."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Sync Open XML SDK data files"
    )
    parser.add_argument(
        "--force", "-f",
        action="store_true",
        help="Force re-download even if up to date"
    )
    parser.add_argument(
        "--check", "-c",
        action="store_true",
        help="Check for updates without downloading"
    )

    args = parser.parse_args()

    if args.check:
        print("Checking for updates...")
        latest = get_latest_commit()
        current = get_current_version()
        print(f"Latest: {latest[:12]}")
        print(f"Current: {current[:12] if current else 'none'}")
        if current == latest:
            print("Up to date!")
        else:
            print("Updates available!")
        return

    sync_data(force=args.force)


if __name__ == "__main__":
    main()
