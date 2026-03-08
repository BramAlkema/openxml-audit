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


def _p95(values: list[float]) -> float:
    if len(values) < 2:
        return values[0]
    return statistics.quantiles(values, n=20)[-1]


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

    total_timings: list[float] = []
    phase_timings: dict[str, list[float]] = {}
    for _ in range(args.iterations):
        start = time.perf_counter()
        _result, phases = validator.validate_with_timings(args.pptx_path)
        measured_total = time.perf_counter() - start
        total_timings.append(measured_total)
        for phase, duration in phases.items():
            phase_timings.setdefault(phase, []).append(duration)

    avg = statistics.mean(total_timings)
    p95 = _p95(total_timings)

    print(f"Iterations: {args.iterations}")
    print(f"Schema: {args.schema}, Semantic: {args.semantic}")
    print(
        f"Avg: {avg:.4f}s, Min: {min(total_timings):.4f}s, "
        f"Max: {max(total_timings):.4f}s, P95: {p95:.4f}s"
    )

    print("\nPhase breakdown (avg/min/max/p95):")
    phase_order = [
        "package_structure",
        "profile_detection",
        "structure",
        "relationships",
        "binary",
        "schema",
        "semantic",
        "specific",
        "total",
    ]
    for phase in phase_order:
        samples = phase_timings.get(phase)
        if not samples:
            continue
        print(
            f"  {phase:18} {statistics.mean(samples):.4f}s / "
            f"{min(samples):.4f}s / {max(samples):.4f}s / {_p95(samples):.4f}s"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
