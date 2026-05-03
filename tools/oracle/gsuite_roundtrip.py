"""Google Workspace roundtrip oracle: does a file survive GSuite import/export?

Spec 031. Sibling of the desktop-app oracles:

  - `tools/oracle/word_repair_oracle.py` — Word for Mac via osascript
  - `tools/oracle/pptx_repair_oracle.py` — PowerPoint for Mac via osascript
  - `tools/oracle/xlsx_repair_oracle.py` — Excel for Mac via osascript
  - `tools/oracle/odf_repair_oracle.py`  — LibreOffice via headless soffice

GSuite is fundamentally different from the desktop oracles: import is
*lossy by design*, so every roundtrip produces a non-zero diff. The
existing `preserved | repaired | ...` outcome vocabulary doesn't carry
useful signal here; instead each observation is classified across a
`LossClass` taxonomy that distinguishes loss (parts removed/degraded)
from normalization (parts GSuite added that weren't in the source).

Per file:

  1. Upload the input to a Drive folder owned by the impersonation
     subject (or to a Shared Drive).
  2. Copy the upload with `mimeType: vnd.google-apps.presentation`
     to import into Slides' native IR.
  3. Export the native Slides file back to .pptx — write the bytes
     to a local working dir.
  4. Hand original + post-GSuite copy to
     `pptx.lab.compare_pptx_packages` for the per-part diff.
  5. Apply the `LossClass` rule-based classifier over the diff.
  6. Roll up into a `GSuiteRoundtripObservation`.
  7. Best-effort delete both Drive files (always, in `finally`).

Spec: `specs/031-gsuite-roundtrip-oracle.md`.
"""

from __future__ import annotations

import argparse
import contextlib
import json
import os
import re
import shutil
import sys
import time
import uuid
import zipfile
from collections.abc import Iterable
from dataclasses import asdict, dataclass, field
from enum import Enum
from pathlib import Path
from typing import Literal

_REPO_ROOT = Path(__file__).resolve().parents[2]
_SRC_DIR = _REPO_ROOT / "src"
if str(_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_SRC_DIR))

from openxml_audit.pptx.lab import compare_pptx_packages  # noqa: E402

_DEFAULT_STAGE_PARENT = Path.home() / "Documents" / ".gsuite_oracle_runs"


class LossClass(str, Enum):
    """Buckets for GSuite roundtrip artifacts.

    Naming policy: stay descriptive. `*_part_changed` / `*_part_removed`
    say what we *observed in the diff* without claiming semantic loss.
    The `*_loss` names (FONT_LOSS, etc., currently unused) are
    *reserved* for future buckets that verify actual loss — e.g., a
    font is removed AND was referenced by a text run; a tableStyles
    file is removed AND the source had `<a:tbl>` shapes using it.
    Until that verification logic exists, file-level signals get
    descriptive names.
    """

    # Descriptive — file-level diff observation.
    THEME_PART_CHANGED = "theme_part_changed"
    MASTER_PART_CHANGED = "master_part_changed"
    STYLE_PART_REMOVED = "style_part_removed"
    FONT_PART_REMOVED = "font_part_removed"
    SLIDE_PART_CHANGED = "slide_part_changed"
    METADATA_CHURN = "metadata_churn"
    MEDIA_RE_ENCODED = "media_re_encoded"
    STRUCTURAL_NORMALIZATION = "structural_normalization"
    DEFAULTS_INLINED = "defaults_inlined"

    # Reserved — verified semantic loss. Not currently fired by any rule;
    # future buckets land here once usage detection is implemented.
    CONTENT_CHANGED = "content_changed"

    # Catch-all for non-empty diffs that didn't match any predicate.
    UNMAPPED = "unmapped"


_OUTCOMES = Literal[
    "preserved",          # zero diff (rare for GSuite — usually impossible)
    "lossy_conversion",   # GSuite returned a result; diff is non-zero
    "upload_failed",
    "convert_failed",
    "export_failed",
    "auth_failed",
    "missing_output",
]


@dataclass
class GSuiteRoundtripObservation:
    """Per-file outcome of a GSuite roundtrip."""

    source_relpath: str
    outcome: _OUTCOMES
    duration_seconds: float
    target_format: str = "pptx"
    subject: str | None = None  # the impersonated Workspace user
    changed_parts: list[str] = field(default_factory=list)
    added_parts: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    loss: list[str] = field(default_factory=list)  # sorted LossClass values
    size_in: int = 0
    size_out: int = 0
    diff_dir: str | None = None
    notes: list[str] = field(default_factory=list)


# --- Loss classifier --------------------------------------------------------
#
# Two layers:
#   1. classify_loss(...) — list-based rules over part names. Cheap, pure,
#      no I/O. Catches structural patterns (theme_loss, master_loss, ...).
#   2. classify_xml_loss(...) — content-aware rules that read slide XML
#      bytes from base/head packages. More expensive but needed for
#      signals that are only visible in element-level diffs (defaults_inlined).
#
# observe() unions the two sets. Tests can exercise either layer in isolation.


# Empty placeholder elements: their presence in the SOURCE means the slide
# inherits values from its layout/master. GSuite's export inlines resolved
# values, so these self-closing tags disappear from the roundtripped slide.
# (We only count self-closing forms — `<p:spPr/>` not `<p:spPr></p:spPr>` —
# because the open/close form already implies content.)
_INHERITED_EMPTY_TAG_RE = re.compile(
    r"<(?:p:spPr|p:cNvSpPr|a:bodyPr|a:pPr|a:xfrm)\s*/>"
)
# Threshold: at least this many empties in the source must vanish in the
# head before we call it inlining (avoids firing on incidental cleanup).
_DEFAULTS_INLINED_MIN_DROP = 2


def _is_pptx_metadata(part: str) -> bool:
    return part.startswith("docProps/") or part == "ppt/viewProps.xml"


def _is_pptx_theme(part: str) -> bool:
    return part.startswith("ppt/theme/")


def _is_pptx_master(part: str) -> bool:
    return part.startswith("ppt/slideMasters/") or part.startswith("ppt/slideLayouts/")


def _is_pptx_notes(part: str) -> bool:
    return part.startswith("ppt/notesMasters/") or part.startswith("ppt/notesSlides/")


def _is_pptx_media(part: str) -> bool:
    return part.startswith("ppt/media/")


def _is_pptx_fonts(part: str) -> bool:
    return part.startswith("ppt/fonts/")


def _is_pptx_styles(part: str) -> bool:
    # tableStyles is the canonical example; future style parts can land here.
    return part in {"ppt/tableStyles.xml"} or part.startswith("ppt/tableStyles")


def _is_pptx_slide(part: str) -> bool:
    return part.startswith("ppt/slides/") and not part.startswith("ppt/slides/_rels/")


def _is_pptx_package_wiring(part: str) -> bool:
    return part in {"[Content_Types].xml", "_rels/.rels"} or part.endswith(".rels")


def classify_loss(
    *,
    changed: Iterable[str],
    added: Iterable[str],
    removed: Iterable[str],
    target_format: str = "pptx",
) -> set[LossClass]:
    """Walk the diff and assign LossClass buckets.

    Phase 1 implements only the pptx ruleset; docx/xlsx land in
    Phase 2 (Spec 031). Multiple classes may fire per roundtrip;
    `unmapped` only fires when the diff is non-empty and nothing
    else matched.
    """
    if target_format != "pptx":
        raise NotImplementedError(
            f"loss classification for {target_format!r} is Phase 2"
        )

    classes: set[LossClass] = set()
    changed_set = set(changed)
    added_set = set(added)
    removed_set = set(removed)
    all_parts = changed_set | added_set | removed_set

    # Package-wiring changes ([Content_Types].xml, *.rels) are a side
    # effect of every other diff and don't carry independent signal —
    # the predicates below match substantive parts only, so wiring-
    # only diffs fall through to UNMAPPED at the bottom.

    if any(_is_pptx_metadata(p) for p in changed_set | removed_set):
        classes.add(LossClass.METADATA_CHURN)
    if any(_is_pptx_theme(p) for p in changed_set | removed_set):
        classes.add(LossClass.THEME_PART_CHANGED)
    if any(_is_pptx_master(p) for p in changed_set | removed_set):
        classes.add(LossClass.MASTER_PART_CHANGED)
    if any(_is_pptx_media(p) for p in changed_set):
        classes.add(LossClass.MEDIA_RE_ENCODED)
    if any(_is_pptx_fonts(p) for p in removed_set):
        classes.add(LossClass.FONT_PART_REMOVED)
    if any(_is_pptx_styles(p) for p in changed_set | removed_set):
        classes.add(LossClass.STYLE_PART_REMOVED)

    # Additions are normalization, not loss — except when GSuite adds
    # parts because it's RE-creating something it dropped (we don't
    # try to detect that pairing in Phase 1).
    if any(_is_pptx_notes(p) for p in added_set):
        classes.add(LossClass.STRUCTURAL_NORMALIZATION)
    # Extra theme variants (theme2.xml, theme3.xml, ...) are GSuite
    # padding the package; theme1 changes go to THEME_PART_CHANGED above.
    extra_themes = [p for p in added_set if _is_pptx_theme(p)]
    if extra_themes:
        classes.add(LossClass.STRUCTURAL_NORMALIZATION)

    if any(_is_pptx_slide(p) for p in changed_set):
        # File-level signal: slide XML changed. We don't claim semantic
        # content loss — that's the reserved CONTENT_CHANGED bucket,
        # which will land once usage-verification logic exists.
        classes.add(LossClass.SLIDE_PART_CHANGED)

    if all_parts and not classes:
        classes.add(LossClass.UNMAPPED)

    return classes


def detect_defaults_inlined(base_xml: bytes, head_xml: bytes) -> bool:
    """Did GSuite expand inheriting empty placeholders into resolved values?

    Counts self-closing inheriting elements in source vs head. A sharp drop
    (>= `_DEFAULTS_INLINED_MIN_DROP`) means GSuite resolved layout/master
    inheritance into inline values during export.

    Caveat: this is a *file-level* signal. We can only observe that
    GSuite's exporter wrote out resolved values, not whether the
    semantic binding to the layout/master is actually broken inside
    Google's IR. The Slides app might still track the inheritance
    internally and only inline-resolve on export. Verifying that
    requires a behavioral oracle (Playwright driving the Slides UI),
    out of scope for the file-level oracle.
    """
    base_text = base_xml.decode("utf-8", errors="replace")
    head_text = head_xml.decode("utf-8", errors="replace")
    base_count = len(_INHERITED_EMPTY_TAG_RE.findall(base_text))
    head_count = len(_INHERITED_EMPTY_TAG_RE.findall(head_text))
    return (base_count - head_count) >= _DEFAULTS_INLINED_MIN_DROP


def classify_xml_loss(
    *,
    base_path: Path,
    head_path: Path,
    target_format: str = "pptx",
) -> set[LossClass]:
    """Content-aware loss detection.

    Opens the two .pptx packages, walks slide XML parts, and runs
    detection rules that need to see actual element content (rather
    than just the diff's part-name list).

    Phase 1 rule set (pptx only):
      - DEFAULTS_INLINED — empty inheriting elements vanish from
        source slides in the roundtripped output.
    """
    if target_format != "pptx":
        return set()

    classes: set[LossClass] = set()

    try:
        base_zip = zipfile.ZipFile(base_path)
        head_zip = zipfile.ZipFile(head_path)
    except zipfile.BadZipFile:
        return set()

    try:
        base_slides = sorted(
            n for n in base_zip.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )
        head_slides = sorted(
            n for n in head_zip.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )

        # Pair slides positionally — works because GSuite preserves slide
        # order. If counts mismatch we check the overlapping prefix only
        # (strict=False is intentional, not an oversight).
        for base_name, head_name in zip(base_slides, head_slides, strict=False):
            base_xml = base_zip.read(base_name)
            head_xml = head_zip.read(head_name)
            if detect_defaults_inlined(base_xml, head_xml):
                classes.add(LossClass.DEFAULTS_INLINED)
                break  # one signal per file is enough
    finally:
        base_zip.close()
        head_zip.close()

    return classes


# --- Orchestrator -----------------------------------------------------------


def _stage_root() -> Path:
    override = os.environ.get("GSUITE_ORACLE_STAGE")
    base = Path(override).expanduser() if override else _DEFAULT_STAGE_PARENT
    base.mkdir(parents=True, exist_ok=True)
    return base


def _stage_input(input_path: Path) -> tuple[Path, Path]:
    work_dir = _stage_root() / f"run-{uuid.uuid4().hex[:8]}"
    work_dir.mkdir(parents=True, exist_ok=False)
    original_copy = work_dir / f"original_{input_path.name}"
    shutil.copy2(input_path, original_copy)
    return work_dir, original_copy


def observe(
    input_pptx: Path,
    *,
    folder_id: str | None = None,
    subject: str | None = None,
    creds_path: Path | str | None = None,
    keep_artifacts: bool = False,
    client: object | None = None,  # injection point for tests
) -> GSuiteRoundtripObservation:
    """Roundtrip one .pptx through GSuite Slides and record the outcome.

    `folder_id` falls back to `GSUITE_ORACLE_FOLDER_ID`; `subject` and
    `creds_path` follow the same pattern as
    `GSuiteClient.from_service_account`. `client` lets tests inject a
    fake; production callers leave it None.
    """
    from openxml_audit.gsuite import (
        PPTX_MIME,
        SLIDES_MIME,
        GSuiteAuthError,
        GSuiteClient,
    )

    input_pptx = input_pptx.resolve()
    if not input_pptx.exists():
        return GSuiteRoundtripObservation(
            source_relpath=input_pptx.name,
            outcome="missing_output",
            duration_seconds=0.0,
            notes=[f"input does not exist: {input_pptx}"],
        )

    resolved_folder = folder_id or os.environ.get("GSUITE_ORACLE_FOLDER_ID")
    if not resolved_folder:
        return GSuiteRoundtripObservation(
            source_relpath=input_pptx.name,
            outcome="auth_failed",
            duration_seconds=0.0,
            subject=subject,
            notes=[
                "no folder_id provided and GSUITE_ORACLE_FOLDER_ID is unset; "
                "create a Drive folder owned by the impersonation subject and "
                "pass its ID. See specs/031-gsuite-roundtrip-oracle.md."
            ],
        )

    started = time.monotonic()
    work_dir, original = _stage_input(input_pptx)
    head = work_dir / f"roundtripped_{input_pptx.name}"
    uploaded_id: str | None = None
    native_id: str | None = None

    try:
        if client is None:
            try:
                client = GSuiteClient.from_service_account(
                    creds_path=creds_path, subject=subject
                )
            except GSuiteAuthError as exc:
                return GSuiteRoundtripObservation(
                    source_relpath=input_pptx.name,
                    outcome="auth_failed",
                    duration_seconds=time.monotonic() - started,
                    subject=subject,
                    notes=[str(exc)],
                )
        active_subject = getattr(client, "subject", subject)

        # 1. Upload
        try:
            uploaded_id = client.upload(
                original, parent_id=resolved_folder, mime_type=PPTX_MIME
            )
        except Exception as exc:
            return GSuiteRoundtripObservation(
                source_relpath=input_pptx.name,
                outcome="upload_failed",
                duration_seconds=time.monotonic() - started,
                subject=active_subject,
                notes=[f"upload failed: {exc}"],
            )

        # 2. Convert
        try:
            native_id = client.convert_to_native(
                uploaded_id,
                target_mime=SLIDES_MIME,
                parent_id=resolved_folder,
                name=input_pptx.stem + "-as-slides",
            )
        except Exception as exc:
            return GSuiteRoundtripObservation(
                source_relpath=input_pptx.name,
                outcome="convert_failed",
                duration_seconds=time.monotonic() - started,
                subject=active_subject,
                notes=[f"convert failed: {exc}"],
            )

        # 3. Export
        try:
            export_bytes = client.export_to_ooxml(native_id, PPTX_MIME)
        except Exception as exc:
            return GSuiteRoundtripObservation(
                source_relpath=input_pptx.name,
                outcome="export_failed",
                duration_seconds=time.monotonic() - started,
                subject=active_subject,
                notes=[f"export failed: {exc}"],
            )
        head.write_bytes(export_bytes)

        # 4. Diff
        compare_dir = work_dir / "compare"
        report = compare_pptx_packages(
            base_path=original, head_path=head, output_dir=compare_dir
        )
        changed = list(report.get("changed_files", []))
        added = list(report.get("added_files", []))
        removed = list(report.get("removed_files", []))

        # 5. Classify — list-based rules over part names + content-aware
        # rules that read slide XML bytes.
        classes = classify_loss(
            changed=changed, added=added, removed=removed, target_format="pptx"
        )
        classes |= classify_xml_loss(
            base_path=original, head_path=head, target_format="pptx"
        )

        return GSuiteRoundtripObservation(
            source_relpath=input_pptx.name,
            outcome="lossy_conversion" if (changed or added or removed) else "preserved",
            duration_seconds=time.monotonic() - started,
            target_format="pptx",
            subject=active_subject,
            changed_parts=changed,
            added_parts=added,
            removed_parts=removed,
            loss=sorted(c.value for c in classes),
            size_in=input_pptx.stat().st_size,
            size_out=head.stat().st_size,
            diff_dir=str(compare_dir) if keep_artifacts else None,
        )
    finally:
        # 6. Cleanup Drive files (best-effort) and local stage dir.
        for fid in (uploaded_id, native_id):
            if fid and client is not None:
                with contextlib.suppress(Exception):
                    client.delete(fid)  # type: ignore[attr-defined]
        if not keep_artifacts:
            shutil.rmtree(work_dir, ignore_errors=True)


def observe_batch(
    inputs: list[Path],
    *,
    folder_id: str | None = None,
    subject: str | None = None,
    creds_path: Path | str | None = None,
    keep_artifacts: bool = False,
) -> list[GSuiteRoundtripObservation]:
    return [
        observe(
            p,
            folder_id=folder_id,
            subject=subject,
            creds_path=creds_path,
            keep_artifacts=keep_artifacts,
        )
        for p in inputs
    ]


def _to_jsonable(observations: list[GSuiteRoundtripObservation]) -> dict:
    loss_counts: dict[str, int] = {}
    for obs in observations:
        for cls in obs.loss:
            loss_counts[cls] = loss_counts.get(cls, 0) + 1

    summary = {
        "total": len(observations),
        "preserved": sum(1 for o in observations if o.outcome == "preserved"),
        "lossy_conversion": sum(1 for o in observations if o.outcome == "lossy_conversion"),
        "upload_failed": sum(1 for o in observations if o.outcome == "upload_failed"),
        "convert_failed": sum(1 for o in observations if o.outcome == "convert_failed"),
        "export_failed": sum(1 for o in observations if o.outcome == "export_failed"),
        "auth_failed": sum(1 for o in observations if o.outcome == "auth_failed"),
        "missing_output": sum(1 for o in observations if o.outcome == "missing_output"),
        "loss_class_counts": loss_counts,
    }
    return {
        "schema_version": 1,
        "engine": "gsuite",
        "observations": [asdict(obs) for obs in observations],
        "summary": summary,
    }


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "input", nargs="+", type=Path,
        help=".pptx files to roundtrip (or directories to walk)",
    )
    p.add_argument(
        "--output", type=Path, default=None,
        help="optional path to write the JSON observation report",
    )
    p.add_argument(
        "--subject", default=None,
        help="Workspace user to impersonate; defaults to GSUITE_ORACLE_SUBJECT env var",
    )
    p.add_argument(
        "--folder-id", default=None,
        help="Drive folder ID for staging; defaults to GSUITE_ORACLE_FOLDER_ID env var",
    )
    p.add_argument(
        "--creds", type=Path, default=None,
        help="path to service account JSON; defaults to GSUITE_ORACLE_CREDS env var "
             "or ~/.config/openxml-audit/google_service_account.json",
    )
    p.add_argument(
        "--keep-artifacts", action="store_true",
        help="leave staging dirs in place for inspection",
    )
    args = p.parse_args()

    inputs: list[Path] = []
    for entry in args.input:
        if entry.is_dir():
            inputs.extend(sorted(entry.rglob("*.pptx")))
        elif entry.is_file() and entry.suffix.lower() == ".pptx":
            inputs.append(entry)

    if not inputs:
        print("no .pptx inputs found", file=sys.stderr)
        return 2

    observations = observe_batch(
        inputs,
        folder_id=args.folder_id,
        subject=args.subject,
        creds_path=args.creds,
        keep_artifacts=args.keep_artifacts,
    )
    report = _to_jsonable(observations)

    if args.output:
        args.output.write_text(json.dumps(report, indent=2))
        print(f"wrote report to {args.output}", file=sys.stderr)
    else:
        print(json.dumps(report, indent=2))

    s = report["summary"]
    print(
        f"\ngsuite-oracle: total={s['total']} "
        f"preserved={s['preserved']} lossy={s['lossy_conversion']} "
        f"upload_failed={s['upload_failed']} convert_failed={s['convert_failed']} "
        f"export_failed={s['export_failed']} auth_failed={s['auth_failed']}",
        file=sys.stderr,
    )
    if s["loss_class_counts"]:
        print(
            "  loss classes: " + ", ".join(
                f"{k}={v}" for k, v in sorted(s["loss_class_counts"].items())
            ),
            file=sys.stderr,
        )
    hard = (
        s["upload_failed"] + s["convert_failed"] + s["export_failed"]
        + s["auth_failed"] + s["missing_output"]
    )
    return 1 if hard > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
