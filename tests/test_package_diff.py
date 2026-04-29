"""Tests for the shared package-diff module (Spec 024, 0.6.8).

`openxml_audit.package_diff` is the format-agnostic per-part diff
extracted from `pptx.lab` so the XLSX, ODF, and Word oracles can
share canonical-c14n + unified-diff machinery instead of each
shipping a hash-only path.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from openxml_audit.package_diff import (
    canonicalize_xml,
    compare_package_parts,
    compare_packages,
    load_package_parts,
    pretty_part_text,
    sanitize_part_name,
    write_part_diff,
)


def _build_pkg(target: Path, parts: dict[str, bytes]) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return target


def test_canonicalize_xml_collapses_whitespace_differences() -> None:
    """Two XML payloads that differ only in whitespace must canonicalize
    to the same bytes — that's what makes the diff fair against
    Office/LibreOffice cosmetic reformatting on save."""
    a = b'<?xml version="1.0"?><root><child a="1" b="2"/></root>'
    b = b'<?xml version="1.0"?>\n<root>\n  <child a="1" b="2"/>\n</root>\n'
    assert canonicalize_xml(a) == canonicalize_xml(b)


def test_canonicalize_xml_falls_back_on_parse_error() -> None:
    """Malformed XML returns stripped raw rather than raising — diff
    callers want a comparable fingerprint, not a hard failure on
    every edge-case file."""
    raw = b"  not really xml  "
    assert canonicalize_xml(raw) == b"not really xml"


def test_load_package_parts_default_filter_picks_xml_and_rels(tmp_path: Path) -> None:
    pkg = _build_pkg(tmp_path / "p.zip", {
        "[Content_Types].xml": b"<x/>",
        "_rels/.rels": b"<r/>",
        "media/img.png": b"\x89PNG fake",
        "data.bin": b"\x00\x01",
    })
    parts = load_package_parts(pkg)
    assert "[Content_Types].xml" in parts
    assert "_rels/.rels" in parts
    assert "media/img.png" not in parts
    assert "data.bin" not in parts
    # Each part has both the raw and canonical form
    assert "raw" in parts["[Content_Types].xml"]
    assert "canonical" in parts["[Content_Types].xml"]


def test_load_package_parts_honors_custom_filter(tmp_path: Path) -> None:
    pkg = _build_pkg(tmp_path / "p.zip", {
        "[Content_Types].xml": b"<x/>",
        "ppt/presentation.xml": b"<p/>",
        "ppt/slides/slide1.xml": b"<s/>",
    })
    parts = load_package_parts(pkg, parts_filter=lambda n: n.startswith("ppt/slides/"))
    assert list(parts) == ["ppt/slides/slide1.xml"]


def test_compare_package_parts_classifies_changed_added_removed() -> None:
    base = {
        "a.xml": {"raw": b"<a/>", "canonical": b"<a/>"},
        "b.xml": {"raw": b"<b/>", "canonical": b"<b/>"},
        "removed.xml": {"raw": b"<x/>", "canonical": b"<x/>"},
    }
    head = {
        "a.xml": {"raw": b"<a/>", "canonical": b"<a/>"},
        "b.xml": {"raw": b"<b CHANGED='yes'/>", "canonical": b"<b CHANGED=\"yes\"/>"},
        "added.xml": {"raw": b"<n/>", "canonical": b"<n/>"},
    }
    diff = compare_package_parts(base, head)
    assert diff["changed"] == ["b.xml"]
    assert diff["added"] == ["added.xml"]
    assert diff["removed"] == ["removed.xml"]


def test_pretty_part_text_pretty_prints_xml() -> None:
    raw = b'<root><a/><b/></root>'
    pretty = pretty_part_text(raw)
    # Pretty-printed: each child on its own line.
    assert "\n" in pretty
    assert "<root>" in pretty
    assert "<a/>" in pretty


def test_pretty_part_text_falls_back_for_malformed() -> None:
    raw = b"not xml"
    assert pretty_part_text(raw) == "not xml"


def test_sanitize_part_name_replaces_slashes() -> None:
    assert sanitize_part_name("ppt/slides/slide1.xml") == "ppt__slides__slide1.xml"
    assert sanitize_part_name("[Content_Types].xml") == "[Content_Types].xml"


def test_write_part_diff_produces_unified_diff(tmp_path: Path) -> None:
    base = b"<root><child>old</child></root>"
    head = b"<root><child>new</child></root>"
    diff_path = tmp_path / "out.diff"
    write_part_diff(
        part_name="word/document.xml",
        base_data=base,
        head_data=head,
        output_path=diff_path,
    )
    text = diff_path.read_text()
    assert "--- base/word/document.xml" in text
    assert "+++ head/word/document.xml" in text
    assert "old" in text
    assert "new" in text


def test_compare_packages_end_to_end_writes_report_and_diffs(tmp_path: Path) -> None:
    base = _build_pkg(tmp_path / "base.zip", {
        "doc.xml": b'<?xml version="1.0"?><doc><p>a</p></doc>',
        "shared.xml": b'<?xml version="1.0"?><x/>',
        "going_away.xml": b'<?xml version="1.0"?><z/>',
    })
    head = _build_pkg(tmp_path / "head.zip", {
        "doc.xml": b'<?xml version="1.0"?><doc><p>b</p></doc>',
        "shared.xml": b'<?xml version="1.0"?><x/>',
        "new.xml": b'<?xml version="1.0"?><n/>',
    })
    out = tmp_path / "compare"
    report = compare_packages(base_path=base, head_path=head, output_dir=out)
    assert report["changed_files"] == ["doc.xml"]
    assert report["added_files"] == ["new.xml"]
    assert report["removed_files"] == ["going_away.xml"]
    # report.json was written
    persisted = json.loads((out / "report.json").read_text())
    assert persisted["changed_files"] == ["doc.xml"]
    # per-part diff for the changed file was emitted
    diff_file = out / "diffs" / "doc.xml.diff"
    assert diff_file.exists()
    assert "<p>a</p>" in diff_file.read_text()
    assert "<p>b</p>" in diff_file.read_text()


def test_compare_packages_canonicalization_ignores_whitespace_only_changes(
    tmp_path: Path,
) -> None:
    """The whole point of the c14n step: if Office/LibreOffice reflows
    whitespace on save, the diff should NOT flag those parts as changed.
    A dual-format check that exercises the canonicalize-before-compare
    path rather than just byte equality."""
    base = _build_pkg(tmp_path / "base.zip", {
        "doc.xml": b'<?xml version="1.0"?><root><a><b/></a></root>',
    })
    head = _build_pkg(tmp_path / "head.zip", {
        # Same XML, but with whitespace and pretty-printing.
        "doc.xml": b'<?xml version="1.0"?>\n<root>\n  <a>\n    <b/>\n  </a>\n</root>\n',
    })
    out = tmp_path / "compare"
    report = compare_packages(base_path=base, head_path=head, output_dir=out)
    assert report["changed_files"] == []
    assert report["added_files"] == []
    assert report["removed_files"] == []
