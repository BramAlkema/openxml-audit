#!/usr/bin/env python3
"""Check for Open XML SDK upstream updates and bump pinned version references.

Usage:
    # Check only (no changes)
    python scripts/check_sdk_update.py --check

    # Update all pinned references from v3.5.1 to v3.6.0
    python scripts/check_sdk_update.py --from v3.5.1 --to v3.6.0

    # Auto-detect latest release and update
    python scripts/check_sdk_update.py --auto
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.request
from pathlib import Path

SDK_REPO = "dotnet/Open-XML-SDK"
SDK_API_URL = f"https://api.github.com/repos/{SDK_REPO}"
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Files containing pinned SDK version references.
PINNED_FILES = [
    ".forgejo/workflows/calibrate-parity.yml",
    ".forgejo/workflows/parity-gate.yml",
    "docs/parity_contract.md",
    "scripts/corpus/run_parity_snapshot.py",
    "scripts/corpus/check_perf_budget.py",
    "scripts/corpus/compare_to_baseline.py",
    "scripts/sync_openxml_data.py",
]

PINNED_PACKAGE_FILES = [
    "scripts/sdk_check/sdk_check.csproj",
    "scripts/sdk_compare/OpenXmlSdkValidator.csproj",
    "tools/parity/dotnet_validator_runner/OpenXmlValidatorRunner.csproj",
]


def fetch_latest_release() -> str:
    """Fetch latest release tag from GitHub."""
    url = f"{SDK_API_URL}/releases/latest"
    req = urllib.request.Request(url)
    req.add_header("Accept", "application/vnd.github.v3+json")
    req.add_header("User-Agent", "openxml-audit")
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode())
        return data["tag_name"]


def find_current_version() -> str | None:
    """Detect the currently pinned SDK version from calibrate-parity.yml."""
    path = PROJECT_ROOT / ".forgejo/workflows/calibrate-parity.yml"
    if not path.exists():
        return None
    text = path.read_text(encoding="utf-8")
    match = re.search(r'default:\s*"(v[\d.]+)"', text)
    return match.group(1) if match else None


def update_file(path: Path, old_version: str, new_version: str) -> int:
    """Replace old version with new in a single file. Returns replacement count."""
    if not path.exists():
        print(f"  SKIP (not found): {path.relative_to(PROJECT_ROOT)}")
        return 0
    text = path.read_text(encoding="utf-8")
    count = text.count(old_version)
    if count == 0:
        return 0
    updated = text.replace(old_version, new_version)
    path.write_text(updated, encoding="utf-8")
    print(f"  {path.relative_to(PROJECT_ROOT)}: {count} replacement(s)")
    return count


def check_only() -> int:
    """Print current vs latest version and exit."""
    current = find_current_version()
    print(f"Current pinned version: {current or 'unknown'}")
    try:
        latest = fetch_latest_release()
    except Exception as exc:
        print(f"Failed to fetch latest release: {exc}", file=sys.stderr)
        return 1
    print(f"Latest upstream release: {latest}")
    if current == latest:
        print("Already up to date.")
        return 0
    print(f"Update available: {current} -> {latest}")
    return 0


def run_update(old_version: str, new_version: str) -> int:
    """Update all pinned references."""
    if old_version == new_version:
        print("Versions are the same, nothing to do.")
        return 0

    print(f"Updating {old_version} -> {new_version}\n")

    total = 0
    for rel_path in PINNED_FILES:
        total += update_file(PROJECT_ROOT / rel_path, old_version, new_version)
    for rel_path in PINNED_PACKAGE_FILES:
        total += update_file(
            PROJECT_ROOT / rel_path,
            old_version.removeprefix("v"),
            new_version.removeprefix("v"),
        )

    print(f"\nDone. {total} replacement(s) across pinned files.")
    print("Next steps:")
    print("  1. Sync data/openxml from the new immutable SDK tag")
    print("  2. Regenerate the corpus manifest and parity baseline; never copy the old baseline")
    print("  3. Run the .NET runners and full Python test suite")
    print("  4. Review parity drift and update waivers only with an explicit rationale")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Check/apply Open XML SDK version bumps.")
    parser.add_argument("--check", action="store_true", help="Check for updates only")
    parser.add_argument("--auto", action="store_true", help="Auto-detect latest and update")
    parser.add_argument("--from", dest="from_version", help="Current version (e.g. v3.5.1)")
    parser.add_argument("--to", dest="to_version", help="Target version (e.g. v3.6.0)")
    args = parser.parse_args()

    if args.check:
        return check_only()

    if args.auto:
        current = find_current_version()
        if not current:
            print("Could not detect current pinned version.", file=sys.stderr)
            return 1
        try:
            latest = fetch_latest_release()
        except Exception as exc:
            print(f"Failed to fetch latest release: {exc}", file=sys.stderr)
            return 1
        if current == latest:
            print(f"Already at latest: {current}")
            return 0
        return run_update(current, latest)

    if args.from_version and args.to_version:
        return run_update(args.from_version, args.to_version)

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
