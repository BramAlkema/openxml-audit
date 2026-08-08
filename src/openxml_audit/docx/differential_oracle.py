"""Cross-editor DOCX preservation matrix.

No editor is treated as the normative source of truth.  Each target output is
compared independently with the original document, then the per-feature
results are aligned to show the portable intersection and target-specific
divergence.
"""

from __future__ import annotations

import hashlib
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from openxml_audit.docx.semantic_snapshot import (
    FEATURES,
    compare_docx_semantics,
    snapshot_docx,
)


@dataclass(frozen=True)
class DifferentialTargetResult:
    target: str
    output_path: str
    package_sha256: str
    semantic_preserved: bool
    changed_features: tuple[str, ...]


def build_docx_differential_report(
    source_path: Path | str,
    targets: dict[str, Path | str],
) -> dict[str, Any]:
    """Compare named editor outputs against one source and align their verdicts."""

    if len(targets) < 2:
        raise ValueError("a differential oracle requires at least two target outputs")
    source = Path(source_path).resolve()
    source_snapshot = snapshot_docx(source)
    target_results: list[DifferentialTargetResult] = []
    comparisons = {}

    for target, raw_path in sorted(targets.items()):
        if not target or not target.strip():
            raise ValueError("target names must be non-empty")
        output = Path(raw_path).resolve()
        snapshot = snapshot_docx(output)
        comparison = compare_docx_semantics(source_snapshot, snapshot)
        comparisons[target] = comparison
        target_results.append(
            DifferentialTargetResult(
                target=target,
                output_path=str(output),
                package_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),
                semantic_preserved=comparison.preserved,
                changed_features=comparison.changed_features,
            )
        )

    matrix = []
    portable_features = []
    divergent_features = []
    for feature in FEATURES:
        per_target = {
            target: {
                "preserved": feature not in comparison.changed_features,
                "feature_sha256": next(
                    item.head_sha256
                    for item in comparison.features
                    if item.feature == feature
                ),
            }
            for target, comparison in sorted(comparisons.items())
        }
        preserved_values = {entry["preserved"] for entry in per_target.values()}
        preserved_by_all = preserved_values == {True}
        if preserved_by_all:
            portable_features.append(feature)
        if len(preserved_values) > 1:
            divergent_features.append(feature)
        matrix.append(
            {
                "feature": feature,
                "source_feature_sha256": next(
                    item.base_sha256
                    for item in next(iter(comparisons.values())).features
                    if item.feature == feature
                ),
                "preserved_by_all": preserved_by_all,
                "targets": per_target,
            }
        )

    return {
        "schema_version": 1,
        "engine": "docx-differential",
        "source": {
            "path": str(source),
            "package_sha256": source_snapshot.source_sha256,
        },
        "targets": [asdict(result) for result in target_results],
        "feature_matrix": matrix,
        "portable_features": portable_features,
        "divergent_features": divergent_features,
        "all_targets_semantically_preserved": all(
            result.semantic_preserved for result in target_results
        ),
    }


__all__ = ["DifferentialTargetResult", "build_docx_differential_report"]
