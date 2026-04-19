from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

from openxml_audit.pptx.lab import (
    _forward_legacy_main,
    _iter_slide_timing_nodes,
    compare_pptx_packages,
)


def _write_test_pptx(path: Path, slide_xml: str) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ppt/slides/slide1.xml", slide_xml)


def test_iter_slide_timing_nodes_includes_direct_and_alternate_content() -> None:
    slide_xml = etree.fromstring(
        """
        <p:sld
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <p:timing/>
          <mc:AlternateContent>
            <mc:Choice Requires="p14">
              <p:timing/>
            </mc:Choice>
            <mc:Fallback>
              <p:timing/>
            </mc:Fallback>
          </mc:AlternateContent>
        </p:sld>
        """
    )

    labels = [label for label, _ in _iter_slide_timing_nodes(slide_xml)]

    assert labels == [
        "direct",
        "alternateContent1.choice1",
        "alternateContent1.fallback",
    ]


def test_compare_pptx_packages_reports_slide_timing_changes(tmp_path: Path) -> None:
    base_pptx = tmp_path / "base.pptx"
    head_pptx = tmp_path / "head.pptx"

    _write_test_pptx(
        base_pptx,
        """
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:timing>
            <p:tnLst>
              <p:par>
                <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="2" dur="500"/>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:tnLst>
          </p:timing>
        </p:sld>
        """,
    )
    _write_test_pptx(
        head_pptx,
        """
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:timing>
            <p:tnLst>
              <p:par>
                <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="2" dur="900"/>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:tnLst>
          </p:timing>
        </p:sld>
        """,
    )

    report = compare_pptx_packages(
        base_path=base_pptx,
        head_path=head_pptx,
        output_dir=tmp_path / "report",
        max_diff_files=5,
    )

    assert report["changed_files"] == ["ppt/slides/slide1.xml"]
    assert report["timing_changes"] == [
        {
            "slide_file": "slide1.xml",
            "status": "changed",
            "labels": ["direct"],
        }
    ]
    assert (tmp_path / "report" / "diffs" / "ppt__slides__slide1.xml.diff").exists()
    assert (tmp_path / "report" / "base" / "slide1" / "direct.timing.normalized.xml").exists()


def test_forward_legacy_main_restores_sys_argv() -> None:
    seen: dict[str, list[str]] = {}
    original_argv = sys.argv[:]

    def fake_main() -> int:
        seen["argv"] = sys.argv[:]
        return 7

    result = _forward_legacy_main(fake_main, ["--", "capture", "--mode", "live"])

    assert result == 7
    assert seen["argv"][1:] == ["capture", "--mode", "live"]
    assert sys.argv == original_argv
