"""Shared scaffold helpers for PPTX oracle decks.

Runtime oracle builders should materialize committed scaffold packages from
``data/pptx_oracle/scaffolds``. A maintenance-only path still exists for
saving through a python-pptx presentation object and then patching authored
timing XML into the package, which is useful when regenerating scaffolds.
"""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from collections.abc import Mapping
from pathlib import Path

from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_INSTALLED_PACKAGE_SCAFFOLD_ROOT = (
    Path(__file__).resolve().parents[1] / "data" / "pptx_oracle" / "scaffolds"
)
_SOURCE_TREE_SCAFFOLD_ROOT = (
    Path(__file__).resolve().parents[3] / "data" / "pptx_oracle" / "scaffolds"
)


def _candidate_scaffold_roots() -> tuple[Path, ...]:
    return tuple(
        root
        for root in (_INSTALLED_PACKAGE_SCAFFOLD_ROOT, _SOURCE_TREE_SCAFFOLD_ROOT)
        if root.is_dir()
    )


def scaffold_root(name: str) -> Path:
    """Return the on-disk directory containing a committed PPTX scaffold."""
    for root in _candidate_scaffold_roots():
        path = root / name
        if path.is_dir():
            return path
    raise FileNotFoundError(f"Unknown PPTX oracle scaffold: {name}")


def materialize_scaffold_package(scaffold_name: str, output_path: Path) -> Path:
    """Build a PPTX by zipping a committed scaffold directory."""
    root = scaffold_root(scaffold_name)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    files = sorted(path for path in root.rglob("*") if path.is_file())
    with zipfile.ZipFile(
        output_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as archive:
        for path in files:
            archive.write(path, path.relative_to(root).as_posix())
    return output_path


def patch_slide_xml_with_timing(slide_xml: bytes, timing_xml: str) -> bytes:
    """Replace the slide's ``<p:timing>`` block with authored timing XML."""
    slide_root = etree.fromstring(slide_xml)
    for existing in slide_root.findall(f"{{{NS_P}}}timing"):
        slide_root.remove(existing)
    slide_root.append(etree.fromstring(timing_xml.encode("utf-8")))
    return etree.tostring(
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
                if (
                    info.filename.startswith("ppt/slides/slide")
                    and info.filename.endswith(".xml")
                ):
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
    "materialize_scaffold_package",
    "patch_slide_xml_with_timing",
    "save_oracle_presentation",
    "scaffold_root",
]
