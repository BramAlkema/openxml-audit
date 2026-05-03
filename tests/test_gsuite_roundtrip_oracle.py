"""Tests for the GSuite roundtrip oracle (Spec 031).

Split into:

  - **always-on**: pure-Python harness logic — LossClass classifier,
    observation shape, dispatcher routing, CLI parsing, mocked Drive
    client. Run on any machine.
  - **GSuite-required**: real-network smoke test. Skipped unless
    `GSUITE_ORACLE_CREDS` (or the default key path) AND
    `GSUITE_ORACLE_SUBJECT` AND `GSUITE_ORACLE_FOLDER_ID` are all set.

The GSuite-required test makes real Drive API calls and modifies a
real Google account — it deliberately keeps the corpus to a single
tiny fixture so a CI run (when configured) costs ~5 seconds and
a few KB of quota.
"""

from __future__ import annotations

import io
import os
import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS = REPO_ROOT / "tools"
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

from oracle.gsuite_roundtrip import (  # noqa: E402
    GSuiteRoundtripObservation,
    LossClass,
    _to_jsonable,
    classify_loss,
    classify_xml_loss,
    detect_defaults_inlined,
    observe,
)

# --- LossClass classifier ---------------------------------------------------


def test_classify_loss_metadata_churn_only():
    classes = classify_loss(
        changed=["docProps/app.xml"], added=[], removed=["docProps/core.xml"],
    )
    assert classes == {LossClass.METADATA_CHURN}


def test_classify_loss_theme_change_vs_extra_theme():
    # Changed theme1 → loss; added theme2 → normalization; both fire.
    classes = classify_loss(
        changed=["ppt/theme/theme1.xml"],
        added=["ppt/theme/theme2.xml"],
        removed=[],
    )
    assert classes == {
        LossClass.THEME_PART_CHANGED,
        LossClass.STRUCTURAL_NORMALIZATION,
    }


def test_classify_loss_master_and_layouts():
    classes = classify_loss(
        changed=[
            "ppt/slideMasters/slideMaster1.xml",
            "ppt/slideLayouts/slideLayout1.xml",
        ],
        added=[],
        removed=[],
    )
    assert classes == {LossClass.MASTER_PART_CHANGED}


def test_classify_loss_notes_added_is_normalization():
    classes = classify_loss(
        changed=[],
        added=[
            "ppt/notesMasters/notesMaster1.xml",
            "ppt/notesSlides/notesSlide1.xml",
        ],
        removed=[],
    )
    assert classes == {LossClass.STRUCTURAL_NORMALIZATION}


def test_classify_loss_table_styles_removed():
    classes = classify_loss(
        changed=[], added=[], removed=["ppt/tableStyles.xml"],
    )
    assert classes == {LossClass.STYLE_PART_REMOVED}


def test_classify_loss_fonts_removed():
    classes = classify_loss(
        changed=[], added=[], removed=["ppt/fonts/font1.fntdata"],
    )
    assert classes == {LossClass.FONT_PART_REMOVED}


def test_classify_loss_media_changed_is_re_encoded():
    classes = classify_loss(
        changed=["ppt/media/image1.png"], added=[], removed=[],
    )
    assert classes == {LossClass.MEDIA_RE_ENCODED}


def test_classify_loss_slide_change_is_slide_part_changed():
    classes = classify_loss(
        changed=["ppt/slides/slide1.xml"], added=[], removed=[],
    )
    assert classes == {LossClass.SLIDE_PART_CHANGED}


def test_classify_loss_package_wiring_only_does_not_loss_classify():
    # If only Content_Types and rels changed, nothing should match
    # the loss buckets — these are package-wiring artifacts. Falls
    # through to UNMAPPED so the diff isn't silently lost.
    classes = classify_loss(
        changed=["[Content_Types].xml", "_rels/.rels"], added=[], removed=[],
    )
    assert classes == {LossClass.UNMAPPED}


def test_classify_loss_empty_diff_yields_empty_set():
    assert classify_loss(changed=[], added=[], removed=[]) == set()


def test_classify_loss_full_presentation1_baseline():
    """Replays the full empirical Presentation1.pptx → GSuite roundtrip
    diff captured during spec development (21 changed, 5 added, 4
    removed). This is the golden-master test: changes here mean
    GSuite or our classifier shifted."""
    changed = [
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/_rels/presentation.xml.rels",
        "ppt/presProps.xml",
        "ppt/presentation.xml",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideLayouts/slideLayout10.xml",
        "ppt/slideLayouts/slideLayout11.xml",
        "ppt/slideLayouts/slideLayout2.xml",
        "ppt/slideLayouts/slideLayout3.xml",
        "ppt/slideLayouts/slideLayout4.xml",
        "ppt/slideLayouts/slideLayout5.xml",
        "ppt/slideLayouts/slideLayout6.xml",
        "ppt/slideLayouts/slideLayout7.xml",
        "ppt/slideLayouts/slideLayout8.xml",
        "ppt/slideLayouts/slideLayout9.xml",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/slides/_rels/slide1.xml.rels",
        "ppt/slides/slide1.xml",
        "ppt/theme/theme1.xml",
    ]
    added = [
        "ppt/notesMasters/_rels/notesMaster1.xml.rels",
        "ppt/notesMasters/notesMaster1.xml",
        "ppt/notesSlides/_rels/notesSlide1.xml.rels",
        "ppt/notesSlides/notesSlide1.xml",
        "ppt/theme/theme2.xml",
    ]
    removed = [
        "docProps/app.xml",
        "docProps/core.xml",
        "ppt/tableStyles.xml",
        "ppt/viewProps.xml",
    ]
    assert len(changed) == 21
    assert len(added) == 5
    assert len(removed) == 4
    classes = classify_loss(changed=changed, added=added, removed=removed)
    assert classes == {
        LossClass.THEME_PART_CHANGED,
        LossClass.MASTER_PART_CHANGED,
        LossClass.STYLE_PART_REMOVED,
        LossClass.METADATA_CHURN,
        LossClass.STRUCTURAL_NORMALIZATION,
        LossClass.SLIDE_PART_CHANGED,
    }


def test_classify_loss_unsupported_format_phase_2():
    with pytest.raises(NotImplementedError):
        classify_loss(changed=[], added=[], removed=[], target_format="docx")


# --- Content-aware classifier (defaults_inlined) ----------------------------


_INHERITING_SLIDE_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="1" name="Title"/><p:cNvSpPr/></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:p><a:pPr/></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:p><a:pPr/></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"""

_INLINED_SLIDE_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="83" name="Google Shape;83;p1"/><p:cNvSpPr txBox="1"/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr anchor="t" lIns="91425"/>
                <a:p><a:pPr indent="0" lvl="0" marL="0"/></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="84" name="Google Shape;84;p1"/><p:cNvSpPr txBox="1"/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr anchor="t" lIns="91425"/>
                <a:p><a:pPr indent="0" lvl="0" marL="0"/></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"""


def test_detect_defaults_inlined_fires_on_canonical_pair():
    """Source slide has 6 inheriting empties (2× spPr/bodyPr/pPr); the
    GSuite-style head has 0. Sharp drop → fire."""
    assert detect_defaults_inlined(_INHERITING_SLIDE_XML, _INLINED_SLIDE_XML) is True


def test_detect_defaults_inlined_no_change_does_not_fire():
    """If both files already have all values inlined, there's nothing to
    detect — drop is 0."""
    assert detect_defaults_inlined(_INLINED_SLIDE_XML, _INLINED_SLIDE_XML) is False


def test_detect_defaults_inlined_below_threshold_does_not_fire():
    """A single empty disappearing isn't enough — could be incidental."""
    base = b"<p:sld><p:spPr/></p:sld>"  # 1 empty
    head = b"<p:sld><p:spPr><a:xfrm/></p:spPr></p:sld>"  # 0 empties
    # Drop of 1 is below the min-drop threshold (2).
    assert detect_defaults_inlined(base, head) is False


def test_detect_defaults_inlined_only_counts_self_closing_form():
    """The open/close form `<p:spPr></p:spPr>` already implies content
    intent and isn't inheriting — we don't count it."""
    base = b"<p:sld><p:spPr></p:spPr><p:spPr></p:spPr></p:sld>"
    head = b"<p:sld><p:spPr><a:xfrm/></p:spPr><p:spPr><a:xfrm/></p:spPr></p:sld>"
    assert detect_defaults_inlined(base, head) is False


def test_classify_xml_loss_pptx_pair(tmp_path):
    """Build a synthetic .pptx pair and run the full content-aware
    classifier over the whole package."""
    base = tmp_path / "base.pptx"
    head = tmp_path / "head.pptx"

    def write_pkg(path: Path, slide_xml: bytes) -> None:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(
                "[Content_Types].xml",
                b'<?xml version="1.0" encoding="UTF-8"?>'
                b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
            )
            z.writestr("ppt/slides/slide1.xml", slide_xml)

    write_pkg(base, _INHERITING_SLIDE_XML)
    write_pkg(head, _INLINED_SLIDE_XML)

    classes = classify_xml_loss(base_path=base, head_path=head)
    assert LossClass.DEFAULTS_INLINED in classes


def test_classify_xml_loss_unsupported_format_returns_empty():
    """docx/xlsx are Phase 2 — silent no-op rather than raising, so
    the orchestrator can union it unconditionally."""
    assert classify_xml_loss(
        base_path=Path("/dev/null"), head_path=Path("/dev/null"),
        target_format="docx",
    ) == set()


# --- Observation shape + JSON rollup ----------------------------------------


def test_observation_dataclass_defaults():
    obs = GSuiteRoundtripObservation(
        source_relpath="x.pptx", outcome="lossy_conversion", duration_seconds=1.0,
    )
    assert obs.target_format == "pptx"
    assert obs.subject is None
    assert obs.loss == []
    assert obs.changed_parts == []


def test_to_jsonable_summary_counts_outcomes_and_loss():
    obs = [
        GSuiteRoundtripObservation(
            source_relpath="a.pptx", outcome="lossy_conversion",
            duration_seconds=1.0, loss=["theme_part_changed", "master_part_changed"],
        ),
        GSuiteRoundtripObservation(
            source_relpath="b.pptx", outcome="lossy_conversion",
            duration_seconds=1.0, loss=["theme_part_changed"],
        ),
        GSuiteRoundtripObservation(
            source_relpath="c.pptx", outcome="auth_failed",
            duration_seconds=0.1,
        ),
    ]
    report = _to_jsonable(obs)
    assert report["engine"] == "gsuite"
    assert report["summary"]["total"] == 3
    assert report["summary"]["lossy_conversion"] == 2
    assert report["summary"]["auth_failed"] == 1
    assert report["summary"]["loss_class_counts"] == {
        "theme_part_changed": 2, "master_part_changed": 1,
    }


# --- Dispatcher wiring ------------------------------------------------------


def test_dispatcher_has_gsuite_engine():
    from openxml_audit.oracle.__main__ import _DISPATCH  # noqa: PLC0415
    assert "gsuite" in _DISPATCH
    assert "google" in _DISPATCH  # alias
    assert _DISPATCH["gsuite"] is _DISPATCH["google"]


# --- Mocked Drive client → observe() exercises the full pipeline ------------


def _build_synthetic_pptx(target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        zf.writestr(
            "ppt/presentation.xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        )
        zf.writestr(
            "ppt/slides/slide1.xml",
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        )
        zf.writestr(
            "docProps/app.xml",
            b'<?xml version="1.0" encoding="UTF-8"?><Properties/>',
        )
    return target


class _FakeClient:
    """Fake GSuiteClient that simulates a lossy GSuite roundtrip
    without any network calls.

    Drops `docProps/app.xml`, adds `ppt/notesMasters/notesMaster1.xml`,
    leaves the rest unchanged.
    """

    subject = "fake@example.com"

    def __init__(self) -> None:
        self.uploaded: list[str] = []
        self.deleted: list[str] = []
        self._stash: bytes | None = None

    def upload(self, path, *, parent_id=None, mime_type=None, name=None):
        self._stash = Path(path).read_bytes()
        self.uploaded.append("upload-id")
        return "upload-id"

    def convert_to_native(self, file_id, *, target_mime, parent_id=None, name=None):
        return "native-id"

    def export_to_ooxml(self, file_id, ooxml_mime):
        # Mutate: drop docProps/app.xml; add ppt/notesMasters/notesMaster1.xml.
        assert self._stash is not None
        out = io.BytesIO()
        with (
            zipfile.ZipFile(io.BytesIO(self._stash), "r") as src,
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst,
        ):
            for info in src.infolist():
                if info.filename == "docProps/app.xml":
                    continue
                dst.writestr(info, src.read(info.filename))
            dst.writestr(
                "ppt/notesMasters/notesMaster1.xml",
                b'<?xml version="1.0" encoding="UTF-8"?>'
                b'<p:notesMaster xmlns:p="http://schemas.openxmlformats.org/'
                b'presentationml/2006/main"/>',
            )
        return out.getvalue()

    def delete(self, file_id):
        self.deleted.append(file_id)
        return True


def test_observe_with_fake_client_classifies_correctly(tmp_path, monkeypatch):
    monkeypatch.setenv("GSUITE_ORACLE_STAGE", str(tmp_path / "stage"))
    src = _build_synthetic_pptx(tmp_path / "synthetic.pptx")
    client = _FakeClient()

    obs = observe(src, folder_id="fake-folder", client=client)

    assert obs.outcome == "lossy_conversion"
    assert obs.subject == "fake@example.com"
    assert "docProps/app.xml" in obs.removed_parts
    assert "ppt/notesMasters/notesMaster1.xml" in obs.added_parts
    # Both METADATA_CHURN (removed app.xml) and STRUCTURAL_NORMALIZATION
    # (added notesMaster) should fire.
    assert "metadata_churn" in obs.loss
    assert "structural_normalization" in obs.loss
    # Both Drive-side files should have been deleted in `finally`.
    assert set(client.deleted) == {"upload-id", "native-id"}


def test_observe_missing_input_returns_missing_output_outcome(tmp_path):
    obs = observe(tmp_path / "does-not-exist.pptx")
    assert obs.outcome == "missing_output"


def test_observe_without_folder_id_fails_fast(tmp_path, monkeypatch):
    """No folder_id → no env var → auth_failed with actionable note.
    Catching this early prevents silently uploading to the user's root
    Drive when GSUITE_ORACLE_FOLDER_ID was forgotten."""
    monkeypatch.delenv("GSUITE_ORACLE_FOLDER_ID", raising=False)
    src = _build_synthetic_pptx(tmp_path / "synthetic.pptx")
    obs = observe(src, folder_id=None, client=_FakeClient())
    assert obs.outcome == "auth_failed"
    assert any("GSUITE_ORACLE_FOLDER_ID" in n for n in obs.notes)


# --- CLI argument parsing ---------------------------------------------------


def test_cli_main_no_inputs_returns_2(tmp_path, monkeypatch, capsys):
    from oracle.gsuite_roundtrip import main as gsuite_main  # noqa: PLC0415

    empty = tmp_path / "empty"
    empty.mkdir()
    monkeypatch.setattr(sys, "argv", ["gsuite_roundtrip.py", str(empty)])
    code = gsuite_main()
    assert code == 2
    err = capsys.readouterr().err
    assert "no .pptx inputs found" in err


# --- GSuite-required smoke test (real network) ------------------------------


def _gsuite_configured() -> bool:
    creds_env = os.environ.get("GSUITE_ORACLE_CREDS")
    creds_default = Path.home() / ".config" / "openxml-audit" / "google_service_account.json"
    has_creds = bool(creds_env) or creds_default.exists()
    return (
        has_creds
        and bool(os.environ.get("GSUITE_ORACLE_SUBJECT"))
        and bool(os.environ.get("GSUITE_ORACLE_FOLDER_ID"))
    )


REQUIRES_GSUITE = pytest.mark.skipif(
    not _gsuite_configured(),
    reason="GSuite oracle requires creds + GSUITE_ORACLE_SUBJECT + GSUITE_ORACLE_FOLDER_ID",
)


@REQUIRES_GSUITE
def test_gsuite_real_roundtrip_smoke():
    """End-to-end test against real GSuite. Uses the Presentation1.pptx
    in the repo root if present; otherwise skips."""
    sample = REPO_ROOT / "Presentation1.pptx"
    if not sample.exists():
        pytest.skip("Presentation1.pptx not in repo root")

    obs = observe(sample)
    assert obs.outcome == "lossy_conversion", obs.notes
    assert obs.size_in > 0
    assert obs.size_out > 0
    # Empirically, GSuite loss on this file should always include at
    # least metadata_churn and structural_normalization. If those
    # disappear, GSuite changed its import behavior — interesting
    # signal worth investigating.
    assert "metadata_churn" in obs.loss
    assert "structural_normalization" in obs.loss
