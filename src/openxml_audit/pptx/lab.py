"""Unified PPTX lab entrypoint for oracle extraction, snapshots, and diffs.

Per-part diff primitives moved to `openxml_audit.package_diff` in 0.6.8 so
the XLSX, ODF, and Word oracles can share the same canonical-c14n + diff
machinery. This module keeps the PPTX-specific public API
(`compare_pptx_packages`, `write_pptx_snapshot`, the timing-tree change
collector) and delegates the format-agnostic parts via thin wrappers.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections.abc import Callable, Sequence
from dataclasses import asdict
from datetime import UTC, datetime
from pathlib import Path

from lxml import etree

from openxml_audit.package_diff import (
    compare_package_parts as _compare_package_parts,
)
from openxml_audit.package_diff import (
    load_package_parts as _load_package_parts,
)
from openxml_audit.package_diff import (
    sanitize_part_name as _sanitize_part_name,
)
from openxml_audit.package_diff import (
    write_part_diff as _write_part_diff,
)
from openxml_audit.pptx.oracle import (
    _build_pattern_index,
    _collect_pptx_paths,
    _extract_deck,
    _normalize_timing_tree,
    _pretty_xml,
    _slugify,
    _write_summary_markdown,
)

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"p": P_NS, "mc": MC_NS}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    snapshot = subparsers.add_parser(
        "snapshot",
        help="Extract readable slide and timing XML views from a PPTX package.",
    )
    snapshot.add_argument("input", type=Path)
    snapshot.add_argument("--output", type=Path, required=True)
    snapshot.set_defaults(func=_cmd_snapshot)

    oracle = subparsers.add_parser(
        "oracle",
        help="Extract timing oracle artifacts from one or more PPTX decks.",
    )
    oracle.add_argument("inputs", nargs="+", type=Path)
    oracle.add_argument("--output", type=Path, required=True)
    oracle.add_argument("--source-name", default="powerpoint-oracle")
    oracle.set_defaults(func=_cmd_oracle)

    diff = subparsers.add_parser(
        "diff",
        help="Extract readable XML/timing views and compare two PPTX packages.",
    )
    diff.add_argument("base", type=Path)
    diff.add_argument("head", type=Path)
    diff.add_argument("--output", type=Path, required=True)
    diff.add_argument(
        "--max-diff-files",
        type=int,
        default=50,
        help="Maximum changed package files to emit unified diffs for.",
    )
    diff.set_defaults(func=_cmd_diff)

    return parser


def _cmd_snapshot(args: argparse.Namespace) -> int:
    write_pptx_snapshot(args.input, args.output)
    return 0


def _cmd_oracle(args: argparse.Namespace) -> int:
    deck_paths = _collect_pptx_paths(args.inputs)
    if not deck_paths:
        raise SystemExit("No .pptx inputs found.")

    output_dir = args.output.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    decks = []
    for deck_path in deck_paths:
        deck_output_dir = output_dir / _slugify(deck_path.stem)
        decks.append(_extract_deck(deck_path, deck_output_dir))

    manifest = {
        "source_name": args.source_name,
        "generated_at": datetime.now(UTC).isoformat(),
        "decks": [asdict(deck) for deck in decks],
        "pattern_index": _build_pattern_index(decks, use_family_signature=False),
        "pattern_family_index": _build_pattern_index(decks, use_family_signature=True),
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    _write_summary_markdown(
        source_name=args.source_name,
        decks=decks,
        pattern_index=manifest["pattern_index"],
        family_index=manifest["pattern_family_index"],
        output_path=output_dir / "README.md",
    )
    return 0


def _cmd_diff(args: argparse.Namespace) -> int:
    compare_pptx_packages(
        base_path=args.base,
        head_path=args.head,
        output_dir=args.output,
        max_diff_files=args.max_diff_files,
    )
    return 0


def _forward_legacy_main(main_func: Callable[[], int | None], args: Sequence[str]) -> int:
    forwarded = list(args)
    if forwarded[:1] == ["--"]:
        forwarded = forwarded[1:]
    previous_argv = sys.argv[:]
    sys.argv = [previous_argv[0], *forwarded]
    try:
        result = main_func()
    finally:
        sys.argv = previous_argv
    return 0 if result is None else int(result)


def compare_pptx_packages(
    *,
    base_path: Path,
    head_path: Path,
    output_dir: Path,
    max_diff_files: int = 50,
) -> dict[str, object]:
    base_path = base_path.resolve()
    head_path = head_path.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    base_snapshot_dir = output_dir / "base"
    head_snapshot_dir = output_dir / "head"
    base_snapshot = write_pptx_snapshot(base_path, base_snapshot_dir)
    head_snapshot = write_pptx_snapshot(head_path, head_snapshot_dir)

    base_parts = _load_package_parts(base_path)
    head_parts = _load_package_parts(head_path)
    part_diff = _compare_package_parts(base_parts, head_parts)

    diff_dir = output_dir / "diffs"
    diff_dir.mkdir(parents=True, exist_ok=True)

    changed_files = part_diff["changed"]
    for part_name in changed_files[: max(0, max_diff_files)]:
        diff_path = diff_dir / f"{_sanitize_part_name(part_name)}.diff"
        _write_part_diff(
            part_name=part_name,
            base_data=base_parts[part_name]["raw"],
            head_data=head_parts[part_name]["raw"],
            output_path=diff_path,
        )

    report = {
        "generated_at": datetime.now(UTC).isoformat(),
        "base": str(base_path),
        "head": str(head_path),
        "base_snapshot": str(base_snapshot_dir),
        "head_snapshot": str(head_snapshot_dir),
        "changed_count": len(changed_files),
        "added_count": len(part_diff["added"]),
        "removed_count": len(part_diff["removed"]),
        "changed_files": changed_files,
        "added_files": part_diff["added"],
        "removed_files": part_diff["removed"],
        "timing_changes": _collect_timing_changes(base_snapshot, head_snapshot),
    }
    (output_dir / "report.json").write_text(
        json.dumps(report, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    _write_diff_report(report=report, output_path=output_dir / "report.md")
    return report


def write_pptx_snapshot(pptx_path: Path, output_dir: Path) -> dict[str, object]:
    pptx_path = pptx_path.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    slides_summary: list[dict[str, object]] = []
    with zipfile.ZipFile(pptx_path) as archive:
        slide_files = sorted(
            name
            for name in archive.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )
        for slide_file in slide_files:
            slide_name = Path(slide_file).name
            slide_bytes = archive.read(slide_file)
            slide_dir = output_dir / slide_name.replace(".xml", "")
            slide_dir.mkdir(parents=True, exist_ok=True)
            slide_views = _extract_slide_views(slide_bytes)
            (slide_dir / "slide.pretty.xml").write_text(
                slide_views["slide_pretty"],
                encoding="utf-8",
            )
            timing_views = {}
            for timing_view in slide_views["timings"]:
                label = timing_view["label"]
                stem = f"{label}.timing"
                (slide_dir / f"{stem}.raw.xml").write_text(
                    timing_view["raw_xml"],
                    encoding="utf-8",
                )
                (slide_dir / f"{stem}.normalized.xml").write_text(
                    timing_view["normalized_xml"],
                    encoding="utf-8",
                )
                timing_views[label] = {
                    "raw_path": f"{slide_name.replace('.xml', '')}/{stem}.raw.xml",
                    "normalized_path": f"{slide_name.replace('.xml', '')}/{stem}.normalized.xml",
                }
            slide_summary = {
                "slide_file": slide_name,
                "timing_labels": [timing["label"] for timing in slide_views["timings"]],
                "timing_views": timing_views,
            }
            (slide_dir / "summary.json").write_text(
                json.dumps(slide_summary, indent=2, sort_keys=True),
                encoding="utf-8",
            )
            slides_summary.append(slide_summary)

    snapshot = {
        "pptx_path": str(pptx_path),
        "snapshot_dir": str(output_dir),
        "generated_at": datetime.now(UTC).isoformat(),
        "slides": slides_summary,
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(snapshot, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    return snapshot


def _extract_slide_views(slide_bytes: bytes) -> dict[str, object]:
    parser = etree.XMLParser(remove_blank_text=True, recover=True)
    slide_xml = etree.fromstring(slide_bytes, parser=parser)
    timings = []
    for label, timing in _iter_slide_timing_nodes(slide_xml):
        timings.append(
            {
                "label": label,
                "raw_xml": _pretty_xml(timing),
                "normalized_xml": _pretty_xml(_normalize_timing_tree(timing)),
            }
        )
    return {
        "slide_pretty": _pretty_xml(slide_xml),
        "timings": timings,
    }


def _iter_slide_timing_nodes(slide_xml: etree._Element) -> list[tuple[str, etree._Element]]:
    nodes: list[tuple[str, etree._Element]] = []

    direct_timing = slide_xml.find("p:timing", namespaces=NS)
    if direct_timing is not None:
        nodes.append(("direct", direct_timing))

    for alt_index, alternate_content in enumerate(
        slide_xml.findall(".//mc:AlternateContent", namespaces=NS),
        start=1,
    ):
        for choice_index, choice in enumerate(
            alternate_content.findall("mc:Choice", namespaces=NS),
            start=1,
        ):
            timing = choice.find("p:timing", namespaces=NS)
            if timing is not None:
                nodes.append((f"alternateContent{alt_index}.choice{choice_index}", timing))
        fallback = alternate_content.find("mc:Fallback", namespaces=NS)
        if fallback is None:
            continue
        timing = fallback.find("p:timing", namespaces=NS)
        if timing is not None:
            nodes.append((f"alternateContent{alt_index}.fallback", timing))

    return nodes


# Per-part diff primitives now live in openxml_audit.package_diff —
# imported above as _load_package_parts / _canonicalize_xml /
# _compare_package_parts / _write_part_diff / _pretty_part_text /
# _sanitize_part_name. Existing PPTX consumers keep working through
# the same private names; the canonical-c14n + diff logic itself is
# now shared with the XLSX / ODF / Word oracles.


def _collect_timing_changes(
    base_snapshot: dict[str, object],
    head_snapshot: dict[str, object],
) -> list[dict[str, object]]:
    base_slides = {slide["slide_file"]: slide for slide in base_snapshot["slides"]}
    head_slides = {slide["slide_file"]: slide for slide in head_snapshot["slides"]}

    changes = []
    for slide_file in sorted(set(base_slides) | set(head_slides)):
        base_slide = base_slides.get(slide_file)
        head_slide = head_slides.get(slide_file)
        if base_slide is None or head_slide is None:
            changes.append(
                {
                    "slide_file": slide_file,
                    "status": "added" if base_slide is None else "removed",
                    "labels": [],
                }
            )
            continue
        base_labels = set(base_slide["timing_labels"])
        head_labels = set(head_slide["timing_labels"])
        changed_labels = sorted(base_labels ^ head_labels)
        shared_labels = sorted(base_labels & head_labels)
        for label in shared_labels:
            base_norm = _read_snapshot_text(
                Path(str(base_snapshot["snapshot_dir"])),
                base_slide["timing_views"][label]["normalized_path"],
            )
            head_norm = _read_snapshot_text(
                Path(str(head_snapshot["snapshot_dir"])),
                head_slide["timing_views"][label]["normalized_path"],
            )
            if base_norm != head_norm:
                changed_labels.append(label)
        if changed_labels:
            changes.append(
                {
                    "slide_file": slide_file,
                    "status": "changed",
                    "labels": sorted(changed_labels),
                }
            )
    return changes


def _read_snapshot_text(snapshot_root: Path, relative_path: str) -> str:
    return (snapshot_root / relative_path).read_text(encoding="utf-8")


def _write_diff_report(*, report: dict[str, object], output_path: Path) -> None:
    lines = [
        "# PPTX Lab Diff",
        "",
        f"Generated: {report['generated_at']}",
        "",
        f"- Base: `{report['base']}`",
        f"- Head: `{report['head']}`",
        f"- Changed package parts: `{report['changed_count']}`",
        f"- Added package parts: `{report['added_count']}`",
        f"- Removed package parts: `{report['removed_count']}`",
        "",
        "## Changed Package Parts",
        "",
    ]
    for name in report["changed_files"]:
        lines.append(f"- `{name}`")
    if not report["changed_files"]:
        lines.append("- None")

    lines.extend(["", "## Added Package Parts", ""])
    for name in report["added_files"]:
        lines.append(f"- `{name}`")
    if not report["added_files"]:
        lines.append("- None")

    lines.extend(["", "## Removed Package Parts", ""])
    for name in report["removed_files"]:
        lines.append(f"- `{name}`")
    if not report["removed_files"]:
        lines.append("- None")

    lines.extend(["", "## Timing Changes", ""])
    for entry in report["timing_changes"]:
        labels = ", ".join(entry["labels"]) if entry["labels"] else "-"
        lines.append(f"- `{entry['slide_file']}`: `{entry['status']}` ({labels})")
    if not report["timing_changes"]:
        lines.append("- None")

    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
