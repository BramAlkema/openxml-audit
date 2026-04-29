"""Per-part diff for OPC and ODF packages (any ZIP-of-XML container).

Format-agnostic: walks the ZIP, identifies XML parts (.xml / .rels by
default; the parts_filter callback can narrow further), canonicalizes
each one via lxml's c14n, and produces a structural report:

  - which parts have identical canonical content
  - which differ (changed)
  - which are added/removed between the two packages
  - per-part text diffs (unified-diff format) under an output dir

Originally extracted from `openxml_audit.pptx.lab` (where it served the
PowerPoint-specific compare_pptx_packages function with timing-tree
diff layered on top). Generalized in 0.6.8 so the XLSX / ODF / Word
roundtrip oracles can use the same shape instead of hash-only deltas.

The "canonicalize" step is critical for fairness: applications like
Excel, Word, PowerPoint, and LibreOffice often reorder attributes or
reflow whitespace during a save without changing semantic content.
A raw byte diff would call those "repaired"; lxml's c14n collapses
the noise so only substantive changes survive into the report.
"""

from __future__ import annotations

import difflib
import json
import zipfile
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path

from lxml import etree


# Default predicate: any .xml or .rels entry inside the zip is a part
# we should canonicalize and diff. Callers with format-specific
# canonical-part lists (e.g. only `xl/worksheets/`, only `ppt/slides/`,
# only ODF top-level XML) can pass a stricter callable.
def _default_parts_filter(name: str) -> bool:
    return name.endswith(".xml") or name.endswith(".rels")


def load_package_parts(
    package_path: Path,
    *,
    parts_filter: Callable[[str], bool] | None = None,
) -> dict[str, dict[str, bytes]]:
    """Load and canonicalize every XML-shaped part in `package_path`.

    Returns `{name: {"raw": bytes, "canonical": bytes}}`. Non-matching
    parts (per `parts_filter`) are skipped silently. Binary parts and
    media never reach this function with the default filter.
    """
    predicate = parts_filter or _default_parts_filter
    parts: dict[str, dict[str, bytes]] = {}
    with zipfile.ZipFile(package_path) as archive:
        for name in archive.namelist():
            if not predicate(name):
                continue
            raw = archive.read(name)
            parts[name] = {
                "raw": raw,
                "canonical": canonicalize_xml(raw),
            }
    return parts


def canonicalize_xml(data: bytes) -> bytes:
    """Return c14n bytes for `data`, or stripped raw on parse error.

    Parse failures fall back to stripped raw rather than raising —
    the diff caller wants a comparable fingerprint, not a hard
    failure on every malformed-but-app-readable file.
    """
    try:
        parser = etree.XMLParser(remove_blank_text=True, recover=True)
        root = etree.fromstring(data, parser=parser)
        return etree.tostring(root, method="c14n")
    except Exception:
        return data.strip()


def compare_package_parts(
    base_parts: dict[str, dict[str, bytes]],
    head_parts: dict[str, dict[str, bytes]],
) -> dict[str, list[str]]:
    """Compute the changed/added/removed sets between two part inventories.

    `changed` is sorted by part name; only parts present in both with
    different canonical bytes are reported. Format-specific code paths
    (e.g. PPTX's timing-tree change collector) layer on top of this.
    """
    base_names = set(base_parts)
    head_names = set(head_parts)
    changed = sorted(
        name
        for name in base_names & head_names
        if base_parts[name]["canonical"] != head_parts[name]["canonical"]
    )
    return {
        "changed": changed,
        "added": sorted(head_names - base_names),
        "removed": sorted(base_names - head_names),
    }


def write_part_diff(
    *,
    part_name: str,
    base_data: bytes,
    head_data: bytes,
    output_path: Path,
) -> None:
    """Write a unified-diff text representation of one part."""
    base_lines = pretty_part_text(base_data).splitlines()
    head_lines = pretty_part_text(head_data).splitlines()
    diff_lines = difflib.unified_diff(
        base_lines,
        head_lines,
        fromfile=f"base/{part_name}",
        tofile=f"head/{part_name}",
        lineterm="",
    )
    output_path.write_text("\n".join(diff_lines) + "\n", encoding="utf-8")


def pretty_part_text(data: bytes) -> str:
    """Pretty-print XML for display in a diff. Falls back to raw text
    on parse error (so the diff still says *something*).

    `recover=True` on the parser means malformed input doesn't raise
    XMLSyntaxError — it returns `None` instead, which would crash
    `etree.tostring`. Both paths fall back to raw decoded text.
    """
    parser = etree.XMLParser(remove_blank_text=True, recover=True)
    try:
        root = etree.fromstring(data, parser=parser)
    except etree.XMLSyntaxError:
        return data.decode("utf-8", errors="replace")
    if root is None:
        return data.decode("utf-8", errors="replace")
    return etree.tostring(root, encoding="unicode", pretty_print=True)


def sanitize_part_name(name: str) -> str:
    """Convert a `path/to/part.xml` part name into a filesystem-safe
    diff-file name. Used to write per-part diffs into a flat directory."""
    return name.replace("/", "__")


def compare_packages(
    *,
    base_path: Path,
    head_path: Path,
    output_dir: Path,
    parts_filter: Callable[[str], bool] | None = None,
    max_diff_files: int = 50,
) -> dict[str, object]:
    """End-to-end compare two OPC/ODF packages and emit a report.

    Writes per-part text diffs into `output_dir/diffs/` (capped at
    `max_diff_files`) and a `report.json` summary. Returns the report
    dict so callers can inspect changed/added/removed and derive their
    own aggregate signal.

    `parts_filter` is the format-specific selector — pass None for
    "any .xml/.rels in the ZIP", or a stricter callable to scope to
    canonical parts of one format (e.g. PPTX's `ppt/slides/` family,
    ODF's top-level `content.xml`/`styles.xml`/etc.).
    """
    base_path = base_path.resolve()
    head_path = head_path.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    base_parts = load_package_parts(base_path, parts_filter=parts_filter)
    head_parts = load_package_parts(head_path, parts_filter=parts_filter)
    part_diff = compare_package_parts(base_parts, head_parts)

    diff_dir = output_dir / "diffs"
    diff_dir.mkdir(parents=True, exist_ok=True)

    changed_files = part_diff["changed"]
    for part_name in changed_files[: max(0, max_diff_files)]:
        diff_path = diff_dir / f"{sanitize_part_name(part_name)}.diff"
        write_part_diff(
            part_name=part_name,
            base_data=base_parts[part_name]["raw"],
            head_data=head_parts[part_name]["raw"],
            output_path=diff_path,
        )

    report = {
        "generated_at": datetime.now(UTC).isoformat(),
        "base": str(base_path),
        "head": str(head_path),
        "changed_count": len(changed_files),
        "added_count": len(part_diff["added"]),
        "removed_count": len(part_diff["removed"]),
        "changed_files": changed_files,
        "added_files": part_diff["added"],
        "removed_files": part_diff["removed"],
    }
    (output_dir / "report.json").write_text(
        json.dumps(report, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    return report


__all__ = [
    "canonicalize_xml",
    "compare_package_parts",
    "compare_packages",
    "load_package_parts",
    "pretty_part_text",
    "sanitize_part_name",
    "write_part_diff",
]
