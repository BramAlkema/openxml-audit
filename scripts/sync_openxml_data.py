#!/usr/bin/env python3
"""Sync Open XML SDK data files from Microsoft's GitHub repository.

Downloads the schema definitions and schematron rules needed to generate
validation constraints from a pinned SDK ref.
"""

from __future__ import annotations

import json
import urllib.parse
import urllib.request
from pathlib import Path

# Configuration
SDK_REPO = "dotnet/Open-XML-SDK"
SDK_REF = "v3.4.1"
SDK_RAW_BASE_URL = f"https://raw.githubusercontent.com/{SDK_REPO}"
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


def _sdk_ref_file() -> Path:
    """Return the file that records the pinned upstream SDK ref."""
    return DATA_DIR / ".sdk_ref"


def _sdk_commit_file() -> Path:
    """Return the file that records the resolved upstream SDK commit."""
    return DATA_DIR / ".sdk_version"


def _sdk_base_url(ref: str) -> str:
    """Build the raw GitHub base URL for a specific SDK ref."""
    return f"{SDK_RAW_BASE_URL}/{ref}"


def get_ref_commit(ref: str = SDK_REF) -> str:
    """Resolve a tag/branch/ref name to a concrete SDK commit hash."""
    quoted_ref = urllib.parse.quote(ref, safe="")
    url = f"{SDK_API_URL}/commits/{quoted_ref}"
    req = urllib.request.Request(url)
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("User-Agent", "openxml-audit")

    with urllib.request.urlopen(req, timeout=30) as response:
        data = json.loads(response.read().decode())
        return data["sha"]


def get_latest_commit() -> str:
    """Backward-compatible alias for the pinned SDK ref commit."""
    return get_ref_commit()


def get_current_version() -> str | None:
    """Get the currently synced SDK commit hash."""
    version_file = _sdk_commit_file()
    if version_file.exists():
        return version_file.read_text(encoding="utf-8").strip()
    return None


def get_current_ref() -> str | None:
    """Get the currently synced SDK tag/branch/ref."""
    ref_file = _sdk_ref_file()
    if ref_file.exists():
        return ref_file.read_text(encoding="utf-8").strip()
    return None


def save_version(commit_hash: str, sdk_ref: str = SDK_REF) -> None:
    """Save the synced SDK ref and resolved commit."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    _sdk_commit_file().write_text(commit_hash, encoding="utf-8")
    _sdk_ref_file().write_text(sdk_ref, encoding="utf-8")


def download_file(url: str, dest: Path) -> None:
    """Download a file from URL to destination."""
    req = urllib.request.Request(url)
    req.add_header("User-Agent", "openxml-audit")

    with urllib.request.urlopen(req, timeout=60) as response:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(response.read())


def get_schema_file_list(ref: str = SDK_REF) -> list[str]:
    """Get list of schema files from the SDK repository for a specific ref."""
    quoted_ref = urllib.parse.quote(ref, safe="")
    url = f"{SDK_API_URL}/contents/data/schemas?ref={quoted_ref}"
    req = urllib.request.Request(url)
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("User-Agent", "openxml-audit")

    with urllib.request.urlopen(req, timeout=30) as response:
        data = json.loads(response.read().decode())
        return sorted(item["name"] for item in data if item["name"].endswith(".json"))


def sync_data(force: bool = False, ref: str = SDK_REF) -> dict[str, int]:
    """Sync data files from Open XML SDK.

    Args:
        force: If True, re-download even if already up to date.
        ref: Git tag/branch/ref to sync from.

    Returns:
        Dictionary with counts of downloaded files.
    """
    print("Checking Open XML SDK repository...")

    latest_commit = get_ref_commit(ref)
    current_version = get_current_version()
    current_ref = get_current_ref()

    print(f"  Pinned ref: {ref}")
    print(f"  Resolved commit: {latest_commit[:12]}")
    print(f"  Current ref: {current_ref or 'unknown'}")
    print(f"  Current commit: {current_version[:12] if current_version else 'none'}")

    if current_version == latest_commit and current_ref in (None, ref) and not force:
        if current_ref != ref:
            save_version(latest_commit, sdk_ref=ref)
        print("Already up to date!")
        return {"schemas": 0, "data_files": 0, "skipped": True}

    # Create data directory
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SCHEMAS_DIR.mkdir(parents=True, exist_ok=True)
    sdk_base_url = _sdk_base_url(ref)

    stats = {"schemas": 0, "data_files": 0, "skipped": False}

    # Download data files
    print("\nDownloading data files...")
    for file_path in DATA_FILES:
        url = f"{sdk_base_url}/{file_path}"
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
    schema_files = get_schema_file_list(ref=ref)
    print(f"  Found {len(schema_files)} schema files")

    for i, filename in enumerate(schema_files, 1):
        url = f"{sdk_base_url}/data/schemas/{filename}"
        dest = SCHEMAS_DIR / filename
        print(f"  [{i}/{len(schema_files)}] {filename}...", end=" ", flush=True)
        try:
            download_file(url, dest)
            print("OK")
            stats["schemas"] += 1
        except Exception as e:
            print(f"FAILED: {e}")

    # Save version
    save_version(latest_commit, sdk_ref=ref)

    print("\nSync complete!")
    print(f"  Schema files: {stats['schemas']}")
    print(f"  Data files: {stats['data_files']}")
    print(f"  SDK ref: {ref}")
    print(f"  SDK commit: {latest_commit[:12]}")

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
    parser.add_argument(
        "--ref",
        default=SDK_REF,
        help=f"SDK tag/branch/ref to sync from (default: {SDK_REF})",
    )

    args = parser.parse_args()

    if args.check:
        print("Checking for updates...")
        latest = get_ref_commit(args.ref)
        current = get_current_version()
        current_ref = get_current_ref()
        print(f"Pinned ref: {args.ref}")
        print(f"Resolved commit: {latest[:12]}")
        print(f"Current ref: {current_ref or 'unknown'}")
        print(f"Current commit: {current[:12] if current else 'none'}")
        if current == latest and current_ref in (None, args.ref):
            print("Up to date!")
        else:
            print("Updates available!")
        return

    sync_data(force=args.force, ref=args.ref)


if __name__ == "__main__":
    main()
