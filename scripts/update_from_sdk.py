#!/usr/bin/env python3
"""Update openxml_audit from Open XML SDK.

This script:
1. Syncs the latest data files from Microsoft's Open XML SDK
2. Verifies the loaders work correctly
3. Reports statistics on loaded schemas and rules
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


def sync_data(force: bool = False) -> bool:
    """Sync data from Open XML SDK."""
    print("=" * 60)
    print("STEP 1: Syncing data from Open XML SDK")
    print("=" * 60)

    from scripts.sync_openxml_data import sync_data as do_sync

    try:
        stats = do_sync(force=force)
        if stats.get("skipped"):
            print("Already up to date.")
        else:
            print(f"\nSynced {stats['schemas']} schema files and {stats['data_files']} data files.")
        return True
    except Exception as e:
        print(f"Error syncing data: {e}")
        return False


def verify_loaders() -> bool:
    """Verify schema and schematron loaders work."""
    print("\n" + "=" * 60)
    print("STEP 2: Verifying loaders")
    print("=" * 60)

    try:
        # Clear any cached registry
        from openxml_audit.codegen import schema_loader, schematron_loader
        schema_loader._registry = None
        schematron_loader._registry = None

        from openxml_audit.codegen import get_schema_registry, get_schematron_registry

        # Test schema registry
        schema_reg = get_schema_registry()
        schema_count = len(schema_reg.list_schemas())
        type_count = schema_reg.count_types()
        element_count = schema_reg.count_elements()

        print(f"\nSchema Registry:")
        print(f"  Namespaces: {schema_count}")
        print(f"  Types: {type_count}")
        print(f"  Elements: {element_count}")

        if schema_count == 0:
            print("ERROR: No schemas loaded!")
            return False

        # Test a known element
        from openxml_audit.codegen import get_element_constraint

        constraint = get_element_constraint(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}sld"
        )
        if constraint is None:
            print("ERROR: Could not load p:sld element!")
            return False
        print(f"  Sample element p:sld: {len(constraint.attributes)} attributes")

        # Test schematron registry
        schematron_reg = get_schematron_registry()
        stats = schematron_reg.get_stats()

        print(f"\nSchematron Registry:")
        print(f"  Total rules: {stats['total']}")
        print(f"  Interpretable: {stats['interpretable']} ({100 * stats['interpretable'] // stats['total']}%)")
        print(f"  Unique contexts: {stats['unique_contexts']}")
        print("  By type:")
        for t, c in stats["by_type"].items():
            print(f"    {t}: {c}")

        return True

    except Exception as e:
        print(f"Error verifying loaders: {e}")
        import traceback
        traceback.print_exc()
        return False


def run_tests() -> bool:
    """Run the test suite."""
    print("\n" + "=" * 60)
    print("STEP 3: Running tests")
    print("=" * 60)

    import subprocess

    result = subprocess.run(
        [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"],
        cwd=Path(__file__).parent.parent,
    )

    return result.returncode == 0


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Update openxml_audit from Open XML SDK"
    )
    parser.add_argument(
        "--force", "-f",
        action="store_true",
        help="Force re-download even if up to date",
    )
    parser.add_argument(
        "--skip-sync",
        action="store_true",
        help="Skip syncing data (verify only)",
    )
    parser.add_argument(
        "--skip-tests",
        action="store_true",
        help="Skip running tests",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Check for updates without downloading",
    )

    args = parser.parse_args()

    if args.check:
        from scripts.sync_openxml_data import get_latest_commit, get_current_version

        print("Checking for updates...")
        try:
            latest = get_latest_commit()
            current = get_current_version()
            print(f"Latest SDK commit: {latest[:12]}")
            print(f"Current version: {current[:12] if current else 'none'}")
            if current == latest:
                print("Up to date!")
            else:
                print("Updates available!")
        except Exception as e:
            print(f"Error checking updates: {e}")
            sys.exit(1)
        return

    success = True

    # Step 1: Sync data
    if not args.skip_sync:
        if not sync_data(force=args.force):
            success = False
    else:
        print("Skipping sync (--skip-sync)")

    # Step 2: Verify loaders
    if success and not verify_loaders():
        success = False

    # Step 3: Run tests
    if not args.skip_tests:
        if success and not run_tests():
            success = False
    else:
        print("\nSkipping tests (--skip-tests)")

    # Summary
    print("\n" + "=" * 60)
    if success:
        print("SUCCESS: All steps completed successfully!")
    else:
        print("FAILURE: Some steps failed. Check output above.")
        sys.exit(1)


if __name__ == "__main__":
    main()
