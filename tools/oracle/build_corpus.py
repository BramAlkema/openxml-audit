"""Build a clean oracle corpus by driving TokenMoulds programmatically.

0.6.9's larger baseline assembled its corpus by globbing TokenMoulds'
filesystem (`reports/visual`, `generated/word-demo`, `scratch/odf`,
…) — which mixed release-quality emitter output with troubleshooting-
session leftovers. The user flagged that mid-run, and Spec 026 0.7.2
deferred a clean rebuild to here.

This driver avoids the corpus-curation problem at its root: instead
of picking files, it *generates* them by calling TokenMoulds'
emitters via their Python API. The output is byte-reproducible (a
function of the chosen brand inputs) and is provably "what TokenMoulds
currently emits."

There's one wrinkle: TokenMoulds' emitters all produce *templates*
(.dotx / .xltx / .potx / .ott / .ots / .otp). Office apps treat
templates differently from documents — Excel opens a `.xltx` by
creating a *new* untitled workbook from the template, which breaks
the oracle's "open the staged file → save → diff" flow (Spec 021
documents this edge case). To get a corpus the four oracles can
roundtrip end-to-end, this script post-processes each template's
bytes to flip its content-type / mimetype to the document variant,
then writes it under the document extension.

The transformation is purely a content-type rename — the underlying
XML is identical to the template form, so any "this is what
TokenMoulds emits" claim about the document corpus is just as true
as if TokenMoulds had emitted it that way directly.

Usage:

    python -m tools.oracle.build_corpus --output-dir data/corpus/tokenmoulds_v0.7.2

Spec 028 (Phase 2 of Spec 026's roadmap to 0.8.0).
"""

from __future__ import annotations

import argparse
import io
import sys
import zipfile
from pathlib import Path
from typing import Literal

# Per-format content-type / mimetype transformations.
#
# OOXML: the [Content_Types].xml lists Override entries by part name.
# We swap the template content type string for the document content
# type string. The transformation is targeted at the main part
# (word/document.xml, xl/workbook.xml, ppt/presentation.xml); other
# parts' content types are unaffected.
#
# ODF: the top-level `mimetype` ZIP entry has the document type as a
# string (no XML). We just rewrite it. META-INF/manifest.xml *also*
# carries the mimetype as `manifest:media-type` on the root file
# entry; we patch that too so consistency holds.

_OOXML_CT_SWAPS = {
    "docx": (
        b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    ),
    "xlsx": (
        b"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ),
    "pptx": (
        b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    ),
}

_ODF_MIMETYPE_SWAPS = {
    "odt": (
        b"application/vnd.oasis.opendocument.text-template",
        b"application/vnd.oasis.opendocument.text",
    ),
    "ods": (
        b"application/vnd.oasis.opendocument.spreadsheet-template",
        b"application/vnd.oasis.opendocument.spreadsheet",
    ),
    "odp": (
        b"application/vnd.oasis.opendocument.presentation-template",
        b"application/vnd.oasis.opendocument.presentation",
    ),
}


def template_to_document(
    package_bytes: bytes,
    *,
    target: Literal["docx", "xlsx", "pptx", "odt", "ods", "odp"],
) -> bytes:
    """Rewrite a template package's content-type / mimetype to its
    document equivalent. Returns new package bytes."""
    if target in _OOXML_CT_SWAPS:
        return _swap_ooxml_content_type(package_bytes, target)
    if target in _ODF_MIMETYPE_SWAPS:
        return _swap_odf_mimetype(package_bytes, target)
    raise ValueError(f"unsupported target format: {target}")


def _swap_ooxml_content_type(package_bytes: bytes, target: str) -> bytes:
    template_ct, document_ct = _OOXML_CT_SWAPS[target]
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(package_bytes)) as src:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    data = data.replace(template_ct, document_ct)
                # Preserve compression + zipinfo metadata
                dst.writestr(info, data)
    return out.getvalue()


def _swap_odf_mimetype(package_bytes: bytes, target: str) -> bytes:
    template_mime, document_mime = _ODF_MIMETYPE_SWAPS[target]
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(package_bytes)) as src:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "mimetype":
                    data = document_mime
                    # mimetype must be the first entry and stored
                    # uncompressed per the ODF spec; preserve that.
                    info_uncompressed = zipfile.ZipInfo("mimetype")
                    dst.writestr(info_uncompressed, data, compress_type=zipfile.ZIP_STORED)
                    continue
                if info.filename == "META-INF/manifest.xml":
                    data = data.replace(template_mime, document_mime)
                dst.writestr(info, data)
    return out.getvalue()


def build_minimal_ir(
    org_id: str = "openxml-audit-corpus",
    locale: str = "US",
):
    """Build a small but valid `DocumentIR` shared across all four
    OOXML/ODF emitters. Same shape used in TokenMoulds' own emitter
    tests.

    The `font_metrics_cache` is needed because TokenMoulds' formulaic
    engine resolves typography tokens (line height, baseline grid)
    from the metrics — without a cache or vendored fonts, the
    AdvancedTokenEngine raises before any emitter runs. The
    committed cache at `data/fonts/cache/metrics.json` in the
    TokenMoulds tree is the canonical reference; resolved relative
    to wherever TokenMoulds is installed.
    """
    # Resolve the font_metrics_cache path relative to TokenMoulds'
    # source tree. The package itself doesn't expose this as a
    # public path; use its module file to locate the project root.
    import tokenmoulds
    from tokenmoulds.dtcg.generator import TokenGenerator
    from tokenmoulds.ir.builder import build_document_ir

    tm_root = Path(tokenmoulds.__file__).resolve().parents[2]
    metrics_cache = tm_root / "data" / "fonts" / "cache" / "metrics.json"
    if not metrics_cache.exists():
        raise FileNotFoundError(
            f"TokenMoulds font metrics cache not found at {metrics_cache}. "
            "The corpus builder needs this to resolve typography tokens."
        )

    tokens = TokenGenerator().generate(
        {
            "font_pair": {"sans": "Inter", "serif": "Roboto Slab"},
            "base_colors": {
                "primary": "#2563EB",
                "secondary": "#DC2626",
            },
            "locale": locale,
            "brand_tone": 0.5,
            "font_metrics_cache": str(metrics_cache),
        }
    )
    return build_document_ir(tokens, org_id, locale)


def emit_word_docx(ir, output: Path) -> int:
    """Emit a .docx via TokenMoulds' WordDocumentEmitter + content-type swap."""
    from tokenmoulds.emitters.word import WordDocumentEmitter

    template_bytes = WordDocumentEmitter(ir).build_package()
    doc_bytes = template_to_document(template_bytes, target="docx")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(doc_bytes)
    return len(doc_bytes)


def emit_excel_xlsx(ir, output: Path) -> int:
    from tokenmoulds.emitters.excel import ExcelEmitter

    template_bytes = ExcelEmitter(ir).build_package()
    doc_bytes = template_to_document(template_bytes, target="xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(doc_bytes)
    return len(doc_bytes)


def emit_powerpoint_pptx(ir, output: Path) -> int:
    from tokenmoulds.emitters.powerpoint import PowerPointEmitter

    template_bytes = PowerPointEmitter(ir).build_package()
    doc_bytes = template_to_document(template_bytes, target="pptx")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(doc_bytes)
    return len(doc_bytes)


def emit_writer_odt(ir, output: Path) -> int:
    from tokenmoulds.emitters.odf import WriterTemplateEmitter

    template_bytes = WriterTemplateEmitter(ir).build_package()
    doc_bytes = template_to_document(template_bytes, target="odt")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(doc_bytes)
    return len(doc_bytes)


def emit_calc_ods(ir, output: Path) -> int:
    from tokenmoulds.emitters.odf import CalcTemplateEmitter

    template_bytes = CalcTemplateEmitter(ir).build_package()
    doc_bytes = template_to_document(template_bytes, target="ods")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(doc_bytes)
    return len(doc_bytes)


def emit_impress_odp(ir, output: Path) -> int:
    from tokenmoulds.emitters.odf import ImpressTemplateEmitter

    template_bytes = ImpressTemplateEmitter(ir).build_package()
    doc_bytes = template_to_document(template_bytes, target="odp")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(doc_bytes)
    return len(doc_bytes)


_ORGS = (
    ("acme", "US"),
    ("globex", "GB"),
)


def build_corpus(output_dir: Path) -> dict[str, list[Path]]:
    """Build the full corpus under `output_dir`. Returns a per-format
    mapping of files written."""
    written: dict[str, list[Path]] = {
        "docx": [],
        "xlsx": [],
        "pptx": [],
        "odt": [],
        "ods": [],
        "odp": [],
    }
    for org_id, locale in _ORGS:
        ir = build_minimal_ir(org_id=org_id, locale=locale)
        suffix = f"{org_id}-{locale.lower()}"

        emitters = (
            ("docx", "word", emit_word_docx),
            ("xlsx", "excel", emit_excel_xlsx),
            ("pptx", "pptx", emit_powerpoint_pptx),
            ("odt", "odf", emit_writer_odt),
            ("ods", "odf", emit_calc_ods),
            ("odp", "odf", emit_impress_odp),
        )
        for ext, format_dir, fn in emitters:
            target = output_dir / format_dir / f"{suffix}.{ext}"
            try:
                size = fn(ir, target)
                written[ext].append(target)
                print(f"  ✓ {target.relative_to(output_dir)} ({size:,} bytes)", file=sys.stderr)
            except Exception as exc:
                print(
                    f"  ✗ {target.relative_to(output_dir)}: {type(exc).__name__}: {exc}",
                    file=sys.stderr,
                )
    return written


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--output-dir",
        type=Path,
        default=Path("data/corpus/tokenmoulds_v0.7.2"),
        help="root of the generated corpus tree",
    )
    args = p.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Building TokenMoulds corpus → {args.output_dir.resolve()}", file=sys.stderr)
    written = build_corpus(args.output_dir)

    total = sum(len(v) for v in written.values())
    print(
        f"\nTotal: {total} files emitted across {sum(1 for v in written.values() if v)} formats",
        file=sys.stderr,
    )
    return 0 if total > 0 else 1


if __name__ == "__main__":
    sys.exit(main())
