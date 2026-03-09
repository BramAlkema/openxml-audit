#!/usr/bin/env python3
"""Benchmark ODF validation runtime."""

from __future__ import annotations

import argparse
import statistics
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))

from openxml_audit import FileFormat, OdfValidator  # noqa: E402


def _p95(values: list[float]) -> float:
    if len(values) < 2:
        return values[0]
    return statistics.quantiles(values, n=20)[-1]


def main() -> int:
    parser = argparse.ArgumentParser(description="Benchmark ODF validation runtime.")
    parser.add_argument("odf_path", type=Path, help="Path to ODF file (ODT/ODS/ODP)")
    parser.add_argument("--iterations", type=int, default=5, help="Number of iterations")
    parser.add_argument(
        "--format",
        choices=["odf_1_2", "odf_1_3"],
        default="odf_1_3",
        help="ODF version to validate against (default: odf_1_3)",
    )
    parser.add_argument("--schema", action="store_true", default=True)
    parser.add_argument("--no-schema", dest="schema", action="store_false")
    parser.add_argument("--semantic", action="store_true", default=True)
    parser.add_argument("--no-semantic", dest="semantic", action="store_false")
    parser.add_argument("--security", action="store_true", default=False)
    parser.add_argument("--no-security", dest="security", action="store_false")
    parser.add_argument(
        "--strict",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Use strict mode (default: True)",
    )
    args = parser.parse_args()

    if not args.odf_path.exists():
        print(f"File not found: {args.odf_path}", file=sys.stderr)
        return 2

    file_format = FileFormat.ODF_1_2 if args.format == "odf_1_2" else FileFormat.ODF_1_3

    validator = OdfValidator(
        file_format=file_format,
        schema_validation=args.schema,
        semantic_validation=args.semantic,
        security_validation=args.security,
        strict=args.strict,
    )

    total_timings: list[float] = []
    phase_timings: dict[str, list[float]] = {}
    last_result = None
    for _ in range(args.iterations):
        start = time.perf_counter()
        result, phases = validator.validate_with_timings(
            args.odf_path,
            include_schema_breakdown=True,
        )
        measured_total = time.perf_counter() - start
        total_timings.append(measured_total)
        last_result = result
        for phase, duration in phases.items():
            phase_timings.setdefault(phase, []).append(duration)

    avg = statistics.mean(total_timings)
    p95 = _p95(total_timings)

    print(f"File: {args.odf_path}")
    print(f"Iterations: {args.iterations}")
    print(f"Format: {file_format.value}, Strict: {args.strict}")
    print(f"Schema: {args.schema}, Semantic: {args.semantic}, Security: {args.security}")
    if last_result is not None:
        print(f"Valid: {last_result.is_valid}, Errors: {last_result.error_count}")
    print(
        f"Avg: {avg:.4f}s, Min: {min(total_timings):.4f}s, "
        f"Max: {max(total_timings):.4f}s, P95: {p95:.4f}s"
    )

    print("\nPhase breakdown (avg/min/max/p95):")
    phase_order = [
        "package_structure",
        "xml_parse",
        "schema",
        "semantic",
        "security",
        "total",
    ]
    extras = sorted(name for name in phase_timings if name not in phase_order)
    phase_order.extend(extras)
    for phase in phase_order:
        samples = phase_timings.get(phase)
        if not samples:
            continue
        print(
            f"  {phase:22} {statistics.mean(samples):.4f}s / "
            f"{min(samples):.4f}s / {max(samples):.4f}s / {_p95(samples):.4f}s"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
