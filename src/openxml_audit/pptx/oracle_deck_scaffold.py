"""Shared scaffold helpers for PPTX oracle decks.

Oracle builders should author the evidence-bearing XML fragments directly and
depend on this module for package persistence. The current implementation saves
through a python-pptx presentation object and then patches authored timing XML
into the package, but the deck builders should not care whether the scaffold
ultimately comes from python-pptx, a checked-in template, or a lower-level XML
package writer.
"""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from collections.abc import Mapping
from pathlib import Path

from lxml import etree as ET

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def patch_slide_xml_with_timing(slide_xml: bytes, timing_xml: str) -> bytes:
    """Replace the slide's ``<p:timing>`` block with authored timing XML."""
    slide_root = ET.fromstring(slide_xml)
    for existing in slide_root.findall(f"{{{NS_P}}}timing"):
        slide_root.remove(existing)
    slide_root.append(ET.fromstring(timing_xml.encode("utf-8")))
    return ET.tostring(
        slide_root,
        encoding="utf-8",
        xml_declaration=True,
        standalone="yes",
    )


def inject_timing_map_into_pptx(
    pptx_path: Path,
    timing_by_slide_number: Mapping[int, str],
    *,
    temp_prefix: str = "pptx-oracle-scaffold-",
) -> None:
    """Patch authored timing XML into an existing PPTX package."""
    temp_dir = Path(tempfile.mkdtemp(prefix=temp_prefix))
    temp_pptx = temp_dir / pptx_path.name
    try:
        with zipfile.ZipFile(pptx_path, "r") as source, zipfile.ZipFile(
            temp_pptx,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as target:
            for info in source.infolist():
                payload = source.read(info.filename)
                if info.filename.startswith("ppt/slides/slide") and info.filename.endswith(".xml"):
                    slide_name = Path(info.filename).stem
                    slide_number = int(slide_name.replace("slide", ""))
                    timing_xml = timing_by_slide_number.get(slide_number)
                    if timing_xml:
                        payload = patch_slide_xml_with_timing(payload, timing_xml)
                target.writestr(info, payload)
        shutil.move(temp_pptx, pptx_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def save_oracle_presentation(
    presentation: object,
    output_path: Path,
    *,
    timing_by_slide_number: Mapping[int, str] | None = None,
    temp_prefix: str = "pptx-oracle-scaffold-",
) -> Path:
    """Persist a scaffolded presentation and inject authored timing fragments."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    save = getattr(presentation, "save", None)
    if not callable(save):  # pragma: no cover - defensive guard
        raise TypeError("presentation must expose a callable save(path) method")
    save(output_path)
    if timing_by_slide_number:
        inject_timing_map_into_pptx(
            output_path,
            timing_by_slide_number,
            temp_prefix=temp_prefix,
        )
    return output_path


__all__ = [
    "NS_P",
    "inject_timing_map_into_pptx",
    "patch_slide_xml_with_timing",
    "save_oracle_presentation",
]
