"""Compare two or more editor-produced DOCX files against one source."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_DIR = REPO_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from openxml_audit.docx.differential_oracle import (  # noqa: E402
    build_docx_differential_report,
)


def _target(value: str) -> tuple[str, Path]:
    name, separator, raw_path = value.partition("=")
    if not separator or not name.strip() or not raw_path.strip():
        raise argparse.ArgumentTypeError("targets must use NAME=PATH")
    return name.strip(), Path(raw_path).expanduser()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="original DOCX")
    parser.add_argument(
        "--target",
        action="append",
        type=_target,
        required=True,
        help="named target output as NAME=PATH; pass at least twice",
    )
    parser.add_argument("--output", type=Path, help="write JSON report")
    args = parser.parse_args()
    targets = dict(args.target)
    if len(targets) < 2:
        parser.error("pass at least two distinct --target NAME=PATH values")
    report = build_docx_differential_report(args.source, targets)
    rendered = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(rendered, end="")
    return 0 if report["all_targets_semantically_preserved"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
