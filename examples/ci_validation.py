#!/usr/bin/env python3
"""Validate all Office files in a directory. Designed for CI pipelines.

Requirements:
    pip install openxml-audit

Usage:
    python ci_validation.py ./output/
    python ci_validation.py ./output/ --format Office2007 --max-errors 50
    python ci_validation.py ./output/ --recursive
    python ci_validation.py ./output/ --parallel 4

Exit codes:
    0 = all files valid
    1 = one or more files invalid
"""

import argparse
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from openxml_audit import OpenXmlValidator, FileFormat

OOXML_EXTENSIONS = {".pptx", ".docx", ".xlsx", ".pptm", ".docm", ".xlsm"}
ODF_EXTENSIONS = {".odt", ".ods", ".odp"}
ALL_EXTENSIONS = OOXML_EXTENSIONS | ODF_EXTENSIONS

FORMAT_MAP = {
    "Office2007": FileFormat.OFFICE_2007,
    "Office2010": FileFormat.OFFICE_2010,
    "Office2013": FileFormat.OFFICE_2013,
    "Office2016": FileFormat.OFFICE_2016,
    "Office2019": FileFormat.OFFICE_2019,
    "Microsoft365": FileFormat.MICROSOFT_365,
}


def find_office_files(directory: Path, recursive: bool = False) -> list[Path]:
    """Find all Office files in a directory."""
    pattern = "**/*" if recursive else "*"
    return sorted(
        p for p in directory.glob(pattern) if p.is_file() and p.suffix.lower() in ALL_EXTENSIONS
    )


def validate_file(
    path: str, file_format: FileFormat, max_errors: int
) -> tuple[str, bool, int, list[str]]:
    """Validate a single file. Designed for use with ProcessPoolExecutor."""
    validator = OpenXmlValidator(file_format=file_format, max_errors=max_errors)
    result = validator.validate(path)
    errors = [e.description for e in result.errors[:5]]
    if result.error_count > 5:
        errors.append(f"... (+{result.error_count - 5} more)")
    return (Path(path).name, result.is_valid, result.error_count, errors)


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate Office files in a directory")
    parser.add_argument("directory", type=Path, help="Directory containing Office files")
    parser.add_argument(
        "--format",
        default="Office2019",
        choices=list(FORMAT_MAP.keys()),
        help="Office version to validate against (default: Office2019)",
    )
    parser.add_argument("--max-errors", type=int, default=100, help="Max errors per file")
    parser.add_argument("--recursive", action="store_true", help="Search subdirectories")
    parser.add_argument(
        "--parallel",
        type=int,
        default=1,
        metavar="N",
        help="Number of parallel workers (default: 1, sequential)",
    )
    args = parser.parse_args()

    if not args.directory.is_dir():
        print(f"Error: {args.directory} is not a directory", file=sys.stderr)
        sys.exit(1)

    files = find_office_files(args.directory, args.recursive)
    if not files:
        print(f"No Office files found in {args.directory}")
        sys.exit(0)

    file_format = FORMAT_MAP[args.format]
    failed = 0

    if args.parallel > 1:
        with ProcessPoolExecutor(max_workers=args.parallel) as pool:
            futures = {
                pool.submit(validate_file, str(p), file_format, args.max_errors): p
                for p in files
            }
            for future in as_completed(futures):
                name, is_valid, error_count, errors = future.result()
                status = "PASS" if is_valid else "FAIL"
                error_info = f" ({error_count} errors)" if not is_valid else ""
                print(f"  [{status}] {name}{error_info}")
                if not is_valid:
                    failed += 1
                    for msg in errors:
                        print(f"         {msg}")
    else:
        validator = OpenXmlValidator(file_format=file_format, max_errors=args.max_errors)
        for path in files:
            result = validator.validate(path)
            status = "PASS" if result.is_valid else "FAIL"
            error_info = f" ({result.error_count} errors)" if not result.is_valid else ""
            print(f"  [{status}] {path.name}{error_info}")
            if not result.is_valid:
                failed += 1
                for error in result.errors[:5]:
                    print(f"         {error.description}")
                if result.error_count > 5:
                    print(f"         ... (+{result.error_count - 5} more)")

    print(f"\n{len(files)} files checked, {len(files) - failed} passed, {failed} failed")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
