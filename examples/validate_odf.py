#!/usr/bin/env python3
"""Validate an ODF file (ODT/ODS/ODP) with openxml-audit.

Requirements:
    pip install openxml-audit

Usage:
    python validate_odf.py document.odt
    python validate_odf.py spreadsheet.ods
    python validate_odf.py presentation.odp
"""

import sys
from pathlib import Path

from openxml_audit.odf import OdfValidator


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python validate_odf.py <file.odt|ods|odp>", file=sys.stderr)
        sys.exit(1)

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"Error: {path} not found", file=sys.stderr)
        sys.exit(1)

    print(f"Validating {path.name}...")
    validator = OdfValidator(
        schema_validation=True,
        semantic_validation=True,
        security_validation=True,
    )
    result = validator.validate(path)

    if result.is_valid:
        print(f"Valid! ({result.warning_count} warnings)")
    else:
        print(f"Invalid: {result.error_count} errors, {result.warning_count} warnings")
        for error in result.errors:
            print(f"  [{error.severity.value}] {error.description}")
            if error.part_uri:
                print(f"    Part: {error.part_uri}")
        sys.exit(1)


if __name__ == "__main__":
    main()
