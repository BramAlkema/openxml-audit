"""Benchmark openxml_audit validation runtime."""

from __future__ import annotations

import argparse
import statistics
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import OpenXmlValidator  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx_path", type=Path, help="Path to PPTX file")
    parser.add_argument("--iterations", type=int, default=5)
    parser.add_argument("--schema", action="store_true", default=True)
    parser.add_argument("--no-schema", dest="schema", action="store_false")
    parser.add_argument("--semantic", action="store_true", default=True)
    parser.add_argument("--no-semantic", dest="semantic", action="store_false")
    args = parser.parse_args()

    if not args.pptx_path.exists():
        print(f"File not found: {args.pptx_path}", file=sys.stderr)
        return 2

    validator = OpenXmlValidator(
        schema_validation=args.schema,
        semantic_validation=args.semantic,
    )

    timings = []
    for _ in range(args.iterations):
        start = time.perf_counter()
        _ = validator.validate(args.pptx_path)
        timings.append(time.perf_counter() - start)

    avg = statistics.mean(timings)
    p95 = statistics.quantiles(timings, n=20)[-1] if len(timings) >= 2 else timings[0]

    print(f"Iterations: {args.iterations}")
    print(f"Schema: {args.schema}, Semantic: {args.semantic}")
    print(f"Avg: {avg:.4f}s, Min: {min(timings):.4f}s, Max: {max(timings):.4f}s, P95: {p95:.4f}s")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
