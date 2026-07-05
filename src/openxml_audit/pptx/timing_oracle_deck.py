"""Build a PowerPoint oracle deck for native timing overrides.

This deck probes the runtime behavior of the PresentationML timing fields we
already emit for SMIL-compatible native mappings:

- ``<p:endCondLst>`` time and click end conditions
- ``repeatDur`` on repeating child ``<p:cTn>`` nodes
- ``restart`` on outer effect containers

Each probe slide includes a shared start trigger and a visual control so the
result can be judged by eye in slideshow mode.
"""

from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from enum import Enum
from itertools import count
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError as exc:  # pragma: no cover - dependency check
    Presentation = None  # type: ignore[assignment]
    RGBColor = None  # type: ignore[assignment]
    MSO_AUTO_SHAPE_TYPE = None  # type: ignore[assignment]
    PP_ALIGN = None  # type: ignore[assignment]
    Inches = None  # type: ignore[assignment]
    Pt = None  # type: ignore[assignment]
    _PPTX_IMPORT_ERROR = exc
else:
    _PPTX_IMPORT_ERROR = None

from lxml import etree

from openxml_audit.pptx.oracle_deck_scaffold import materialize_scaffold_package
from openxml_audit.pptx.oracle_starter_deck import (
    NS_P,
    SLIDE_WIDTH_IN,
    BuildEntry,
    StartCondition,
    _add_oracle_marker,
    _add_textbox,
    _build_anim_motion,
    _build_effect_par,
    p_elem,
    p_sub,
)

logger = logging.getLogger(__name__)

START_LEFT_IN = 2.85
LANE_WIDTH_IN = 6.15
SHORT_HOP_IN = 1.35
RUNNER_WIDTH_IN = 1.0
RUNNER_HEIGHT_IN = 0.48


class TriggerType(str, Enum):
    TIME_OFFSET = "time_offset"
    CLICK = "click"
    ELEMENT_BEGIN = "element_begin"
    ELEMENT_END = "element_end"


@dataclass(frozen=True, slots=True)
class Trigger:
    trigger_type: TriggerType
    delay_seconds: float = 0.0
    target_element_id: str | None = None


def _require_python_pptx() -> None:
    if Presentation is None or Inches is None or Pt is None:
        raise RuntimeError(
            "python-pptx is required to build the timing oracle deck."
        ) from _PPTX_IMPORT_ERROR


def _slide(presentation):
    return presentation.slides.add_slide(presentation.slide_layouts[6])


def _note(slide, title: str, subtitle: str) -> None:
    _add_textbox(
        slide,
        0.55,
        0.3,
        12.2,
        0.42,
        title,
        font_size=20,
        bold=True,
        color=(31, 41, 55),
        align=PP_ALIGN.CENTER,
    )
    _add_textbox(
        slide,
        0.85,
        0.74,
        11.6,
        0.44,
        subtitle,
        font_size=12,
        color=(75, 85, 99),
        align=PP_ALIGN.CENTER,
    )


def _add_track(slide, *, top: float) -> None:
    track = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(START_LEFT_IN),
        Inches(top + 0.17),
        Inches(LANE_WIDTH_IN),
        Inches(0.12),
    )
    track.fill.solid()
    track.fill.fore_color.rgb = RGBColor(226, 232, 240)
    track.line.color.rgb = RGBColor(203, 213, 225)
    track.line.width = Pt(0.5)


def _add_lane_label(slide, *, top: float, text: str) -> None:
    _add_textbox(
        slide,
        0.75,
        top - 0.03,
        1.8,
        0.28,
        text,
        font_size=13,
        bold=True,
        color=(55, 65, 81),
        align=PP_ALIGN.RIGHT,
    )


def _add_runner(
    slide,
    *,
    left: float,
    top: float,
    label: str,
    fill_rgb: tuple[int, int, int],
) -> object:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(RUNNER_WIDTH_IN),
        Inches(RUNNER_HEIGHT_IN),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*fill_rgb)
    shape.line.color.rgb = RGBColor(30, 41, 59)
    shape.line.width = Pt(1.0)
    text_frame = shape.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = label
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(15, 23, 42)
    return shape


def _add_button(
    slide,
    *,
    left: float,
    top: float,
    width: float,
    text: str,
    fill_rgb: tuple[int, int, int],
) -> object:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(0.62),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*fill_rgb)
    shape.line.color.rgb = RGBColor(30, 41, 59)
    shape.line.width = Pt(1.25)
    text_frame = shape.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    return shape


def _motion_path(distance_in: float) -> str:
    return f"M 0 0 L {distance_in / SLIDE_WIDTH_IN:.6f} 0.000000 E"


def _max_shape_id(slide) -> int:
    return max(shape.shape_id for shape in slide.shapes)


def _motion_effect(
    *,
    id_counter,
    target_shape_id: int,
    duration_ms: int,
    grp_id: int,
    distance_in: float,
    start_conditions: tuple[StartCondition, ...] = (StartCondition(delay_ms=0),),
    restart: str | None = None,
    repeat_count: str | int | None = None,
    repeat_duration_ms: int | None = None,
    end_triggers: list[Trigger] | None = None,
) -> object:
    motion = _build_anim_motion(
        id_counter=id_counter,
        target_shape_id=target_shape_id,
        duration_ms=duration_ms,
        path=_motion_path(distance_in),
    )
    motion_ctn = motion.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}cBhvr/"
        "{http://schemas.openxmlformats.org/presentationml/2006/main}cTn"
    )
    if repeat_count is not None and motion_ctn is not None:
        motion_ctn.set("repeatCount", str(repeat_count))

    par = _build_effect_par(
        id_counter=id_counter,
        duration_ms=duration_ms,
        node_type="clickEffect",
        preset_id=31,
        preset_class="path",
        grp_id=grp_id,
        child_elements=[motion],
        start_conditions=start_conditions,
    )
    _apply_native_timing_overrides(
        par=par,
        repeat_duration_ms=repeat_duration_ms,
        restart=restart,
        end_triggers=end_triggers,
        default_target_shape=str(target_shape_id),
    )
    return par


def _apply_native_timing_overrides(
    *,
    par: etree._Element,
    repeat_duration_ms: int | None = None,
    restart: str | None = None,
    end_triggers: list[Trigger] | None = None,
    default_target_shape: str | None = None,
) -> None:
    ctn = par.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cTn")
    if ctn is None:
        return

    if restart in {"always", "whenNotActive", "never"}:
        ctn.set("restart", restart)

    if repeat_duration_ms is not None:
        repeat_duration = str(max(1, repeat_duration_ms))
        targets = [node for node in par.iter(f"{{{NS_P}}}cTn") if node.get("repeatCount")]
        for target in (targets or [ctn]):
            target.set("repeatDur", repeat_duration)

    if not end_triggers:
        return

    end_cond_lst = ctn.find(f"{{{NS_P}}}endCondLst")
    if end_cond_lst is None:
        # CT_TLCommonTimeNodeData orders endCondLst after stCondLst and
        # before childTnLst; appending at the end is not schema-valid.
        end_cond_lst = etree.Element(f"{{{NS_P}}}endCondLst")
        st_cond_lst = ctn.find(f"{{{NS_P}}}stCondLst")
        insert_at = ctn.index(st_cond_lst) + 1 if st_cond_lst is not None else 0
        ctn.insert(insert_at, end_cond_lst)
    _append_end_conditions(
        end_cond_lst=end_cond_lst,
        end_triggers=end_triggers,
        default_target_shape=default_target_shape,
    )


def _append_end_conditions(
    *,
    end_cond_lst: etree._Element,
    end_triggers: list[Trigger],
    default_target_shape: str | None,
) -> None:
    created = 0
    for trigger in end_triggers:
        delay_ms = max(0, int(round(trigger.delay_seconds * 1000)))
        if trigger.trigger_type == TriggerType.TIME_OFFSET:
            p_sub(end_cond_lst, "cond", delay=str(delay_ms))
            created += 1
            continue

        if trigger.trigger_type == TriggerType.CLICK:
            cond = p_sub(end_cond_lst, "cond", evt="onClick", delay=str(delay_ms))
            target_shape = trigger.target_element_id or default_target_shape
            if target_shape:
                tgt_el = p_sub(cond, "tgtEl")
                p_sub(tgt_el, "spTgt", spid=target_shape)
            created += 1
            continue

        if trigger.trigger_type in {TriggerType.ELEMENT_BEGIN, TriggerType.ELEMENT_END}:
            cond = p_sub(
                end_cond_lst,
                "cond",
                evt="onBegin" if trigger.trigger_type == TriggerType.ELEMENT_BEGIN else "onEnd",
                delay=str(delay_ms),
            )
            if trigger.target_element_id:
                tgt_el = p_sub(cond, "tgtEl")
                p_sub(tgt_el, "spTgt", spid=trigger.target_element_id)
            created += 1

    if created == 0:
        parent = end_cond_lst.getparent()
        if parent is not None:
            parent.remove(end_cond_lst)


def _build_auto_timing_xml(
    *,
    start_id: int,
    main_effect_pars: tuple[object, ...],
    build_entries: tuple[BuildEntry, ...],
) -> str:
    id_counter = count(start_id)
    timing = p_elem("timing")
    tn_lst = p_sub(timing, "tnLst")
    root_par = p_sub(tn_lst, "par")
    root_ctn = p_sub(
        root_par,
        "cTn",
        id=str(next(id_counter)),
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )
    root_children = p_sub(root_ctn, "childTnLst")
    seq = p_sub(root_children, "seq", concurrent="1", nextAc="seek")
    seq_ctn = p_sub(
        seq,
        "cTn",
        id=str(next(id_counter)),
        dur="indefinite",
        nodeType="mainSeq",
    )
    seq_children = p_sub(seq_ctn, "childTnLst")
    outer_par = p_sub(seq_children, "par")
    outer_ctn = p_sub(outer_par, "cTn", id=str(next(id_counter)), fill="hold")
    outer_st = p_sub(outer_ctn, "stCondLst")
    p_sub(outer_st, "cond", delay="0")
    outer_children = p_sub(outer_ctn, "childTnLst")
    inner_par = p_sub(outer_children, "par")
    inner_ctn = p_sub(inner_par, "cTn", id=str(next(id_counter)), fill="hold")
    inner_st = p_sub(inner_ctn, "stCondLst")
    p_sub(inner_st, "cond", delay="0")
    inner_children = p_sub(inner_ctn, "childTnLst")
    for effect_par in main_effect_pars:
        inner_children.append(effect_par)

    prev_cond_lst = p_sub(seq, "prevCondLst")
    prev_cond = p_sub(prev_cond_lst, "cond", evt="onPrev", delay="0")
    p_sub(p_sub(prev_cond, "tgtEl"), "sldTgt")
    next_cond_lst = p_sub(seq, "nextCondLst")
    next_cond = p_sub(next_cond_lst, "cond", evt="onNext", delay="0")
    p_sub(p_sub(next_cond, "tgtEl"), "sldTgt")

    if build_entries:
        bld_lst = p_sub(timing, "bldLst")
        for entry in build_entries:
            p_sub(
                bld_lst,
                "bldP",
                spid=str(entry.shape_id),
                grpId=entry.grp_id,
                animBg="1",
            )

    return etree.tostring(timing, encoding="unicode")


def _title_slide(presentation) -> str | None:
    slide = _slide(presentation)
    _note(
        slide,
        "PowerPoint Timing Oracle",
        "Probe deck for native end conditions, repeat duration, and restart behavior.",
    )
    bullets = [
        "Slide 2: end=2s against a full-duration control.",
        "Slide 3: stop-on-click against a control that ignores STOP.",
        "Slide 4: repeatDur=2.5s against an uncapped loop.",
        "Slide 5: restart overlap with always / whenNotActive / never.",
        "Slide 6: restart after idle with the same three modes.",
    ]
    top = 1.85
    for bullet in bullets:
        _add_textbox(
            slide,
            1.15,
            top,
            11.2,
            0.34,
            f"- {bullet}",
            font_size=16,
            color=(31, 41, 55),
        )
        top += 0.5
    _add_textbox(
        slide,
        1.15,
        5.2,
        11.2,
        0.7,
        (
            "Run in slideshow mode. Use the buttons on each slide instead of "
            "blank-area clicks so the trigger path stays explicit."
        ),
        font_size=15,
        color=(75, 85, 99),
        align=PP_ALIGN.CENTER,
    )
    return None


def _end_offset_slide(presentation) -> str:
    slide = _slide(presentation)
    _note(
        slide,
        "Time End Condition",
        (
            "On slide entry, the control should finish the lane; the lower runner "
            "should stop at the 2s marker."
        ),
    )

    top_lane = 2.15
    bottom_lane = 3.65
    _add_lane_label(slide, top=top_lane, text="Control")
    _add_lane_label(slide, top=bottom_lane, text="end=2s")
    _add_track(slide, top=top_lane)
    _add_track(slide, top=bottom_lane)
    _add_oracle_marker(
        slide,
        left=START_LEFT_IN + (LANE_WIDTH_IN * 0.4),
        top=1.82,
        height=2.5,
        label="2s",
    )
    _add_oracle_marker(slide, left=START_LEFT_IN + LANE_WIDTH_IN, top=1.82, height=2.5, label="5s")

    control = _add_runner(
        slide,
        left=START_LEFT_IN,
        top=top_lane,
        label="control",
        fill_rgb=(147, 197, 253),
    )
    probe = _add_runner(
        slide,
        left=START_LEFT_IN,
        top=bottom_lane,
        label="end=2s",
        fill_rgb=(251, 191, 36),
    )

    max_shape_id = _max_shape_id(slide)
    id_counter = count(max_shape_id + 20)
    control_effect = _motion_effect(
        id_counter=id_counter,
        target_shape_id=control.shape_id,
        duration_ms=5000,
        grp_id=1,
        distance_in=LANE_WIDTH_IN,
    )
    probe_effect = _motion_effect(
        id_counter=id_counter,
        target_shape_id=probe.shape_id,
        duration_ms=5000,
        grp_id=2,
        distance_in=LANE_WIDTH_IN,
        end_triggers=[
            Trigger(
                trigger_type=TriggerType.TIME_OFFSET,
                delay_seconds=2.0,
            )
        ],
    )
    return _build_auto_timing_xml(
        start_id=max_shape_id + 1,
        main_effect_pars=(control_effect, probe_effect),
        build_entries=(
            BuildEntry(shape_id=control.shape_id, grp_id=1),
            BuildEntry(shape_id=probe.shape_id, grp_id=2),
        ),
    )


def _end_click_slide(presentation) -> str:
    slide = _slide(presentation)
    _note(
        slide,
        "Click End Condition",
        (
            "On slide entry, both runners should start. Click STOP while they are "
            "moving; the lower runner should stop immediately while the control "
            "keeps going."
        ),
    )
    stop_button = _add_button(
        slide,
        left=10.85,
        top=6.15,
        width=1.45,
        text="STOP",
        fill_rgb=(220, 38, 38),
    )

    top_lane = 2.15
    bottom_lane = 3.65
    _add_lane_label(slide, top=top_lane, text="Control")
    _add_lane_label(slide, top=bottom_lane, text="stop.click")
    _add_track(slide, top=top_lane)
    _add_track(slide, top=bottom_lane)
    _add_oracle_marker(
        slide,
        left=START_LEFT_IN + LANE_WIDTH_IN,
        top=1.82,
        height=2.5,
        label="finish",
    )

    control = _add_runner(
        slide,
        left=START_LEFT_IN,
        top=top_lane,
        label="control",
        fill_rgb=(147, 197, 253),
    )
    probe = _add_runner(
        slide,
        left=START_LEFT_IN,
        top=bottom_lane,
        label="stop.click",
        fill_rgb=(251, 191, 36),
    )

    max_shape_id = _max_shape_id(slide)
    id_counter = count(max_shape_id + 20)
    control_effect = _motion_effect(
        id_counter=id_counter,
        target_shape_id=control.shape_id,
        duration_ms=5000,
        grp_id=1,
        distance_in=LANE_WIDTH_IN,
    )
    probe_effect = _motion_effect(
        id_counter=id_counter,
        target_shape_id=probe.shape_id,
        duration_ms=5000,
        grp_id=2,
        distance_in=LANE_WIDTH_IN,
        end_triggers=[
            Trigger(
                trigger_type=TriggerType.CLICK,
                target_element_id=str(stop_button.shape_id),
                delay_seconds=0.0,
            )
        ],
    )
    return _build_auto_timing_xml(
        start_id=max_shape_id + 1,
        main_effect_pars=(control_effect, probe_effect),
        build_entries=(
            BuildEntry(shape_id=control.shape_id, grp_id=1),
            BuildEntry(shape_id=probe.shape_id, grp_id=2),
        ),
    )


def _repeat_duration_slide(presentation) -> str:
    slide = _slide(presentation)
    _note(
        slide,
        "Repeat Duration",
        (
            "On slide entry, both runners should loop the short hop; the lower "
            "runner should stop after about 2.5 seconds."
        ),
    )

    top_lane = 2.15
    bottom_lane = 3.65
    _add_lane_label(slide, top=top_lane, text="control loop")
    _add_lane_label(slide, top=bottom_lane, text="repeatDur=2.5s")
    _add_track(slide, top=top_lane)
    _add_track(slide, top=bottom_lane)
    _add_oracle_marker(slide, left=START_LEFT_IN + SHORT_HOP_IN, top=1.82, height=2.5, label="hop")

    control = _add_runner(
        slide,
        left=START_LEFT_IN,
        top=top_lane,
        label="loop",
        fill_rgb=(167, 243, 208),
    )
    probe = _add_runner(
        slide,
        left=START_LEFT_IN,
        top=bottom_lane,
        label="2.5s cap",
        fill_rgb=(251, 191, 36),
    )

    max_shape_id = _max_shape_id(slide)
    id_counter = count(max_shape_id + 20)
    control_effect = _motion_effect(
        id_counter=id_counter,
        target_shape_id=control.shape_id,
        duration_ms=800,
        grp_id=1,
        distance_in=SHORT_HOP_IN,
        repeat_count="indefinite",
    )
    probe_effect = _motion_effect(
        id_counter=id_counter,
        target_shape_id=probe.shape_id,
        duration_ms=800,
        grp_id=2,
        distance_in=SHORT_HOP_IN,
        repeat_count="indefinite",
        repeat_duration_ms=2500,
    )
    return _build_auto_timing_xml(
        start_id=max_shape_id + 1,
        main_effect_pars=(control_effect, probe_effect),
        build_entries=(
            BuildEntry(shape_id=control.shape_id, grp_id=1),
            BuildEntry(shape_id=probe.shape_id, grp_id=2),
        ),
    )


def _restart_overlap_slide(presentation) -> str:
    slide = _slide(presentation)
    _note(
        slide,
        "Restart During Active Playback",
        (
            "On slide entry, a second begin fires at 1s. Only 'always' should "
            "restart mid-run; the other two should keep their first run."
        ),
    )

    lane_tops = (1.95, 3.15, 4.35)
    labels = ("always", "whenNotActive", "never")
    fills = ((147, 197, 253), (167, 243, 208), (251, 191, 36))
    runners = []
    for top, label, fill in zip(lane_tops, labels, fills, strict=True):
        _add_lane_label(slide, top=top, text=label)
        _add_track(slide, top=top)
        runners.append(
            _add_runner(
                slide,
                left=START_LEFT_IN,
                top=top,
                label=label,
                fill_rgb=fill,
            )
        )
    _add_oracle_marker(
        slide,
        left=START_LEFT_IN + (LANE_WIDTH_IN * 0.25),
        top=1.62,
        height=3.35,
        label="1s",
    )

    start_conditions = (
        StartCondition(delay_ms=0),
        StartCondition(delay_ms=1000),
    )
    max_shape_id = _max_shape_id(slide)
    id_counter = count(max_shape_id + 20)
    effects = tuple(
        _motion_effect(
            id_counter=id_counter,
            target_shape_id=runner.shape_id,
            duration_ms=4000,
            grp_id=index,
            distance_in=LANE_WIDTH_IN,
            start_conditions=start_conditions,
            restart=restart_mode,
        )
        for index, (runner, restart_mode) in enumerate(
            zip(runners, ("always", "whenNotActive", "never"), strict=True),
            start=1,
        )
    )
    return _build_auto_timing_xml(
        start_id=max_shape_id + 1,
        main_effect_pars=effects,
        build_entries=tuple(
            BuildEntry(shape_id=runner.shape_id, grp_id=index)
            for index, runner in enumerate(runners, start=1)
        ),
    )


def _restart_idle_slide(presentation) -> str:
    slide = _slide(presentation)
    _note(
        slide,
        "Restart After Idle",
        (
            "On slide entry, a second begin fires at 5s after the first run is "
            "over. 'always' and 'whenNotActive' should run again; 'never' should "
            "stay put."
        ),
    )

    lane_tops = (1.95, 3.15, 4.35)
    labels = ("always", "whenNotActive", "never")
    fills = ((147, 197, 253), (167, 243, 208), (251, 191, 36))
    runners = []
    for top, label, fill in zip(lane_tops, labels, fills, strict=True):
        _add_lane_label(slide, top=top, text=label)
        _add_track(slide, top=top)
        runners.append(
            _add_runner(
                slide,
                left=START_LEFT_IN,
                top=top,
                label=label,
                fill_rgb=fill,
            )
        )
    _add_oracle_marker(
        slide,
        left=START_LEFT_IN + LANE_WIDTH_IN,
        top=1.62,
        height=3.35,
        label="finish",
    )

    start_conditions = (
        StartCondition(delay_ms=0),
        StartCondition(delay_ms=5000),
    )
    max_shape_id = _max_shape_id(slide)
    id_counter = count(max_shape_id + 20)
    effects = tuple(
        _motion_effect(
            id_counter=id_counter,
            target_shape_id=runner.shape_id,
            duration_ms=2500,
            grp_id=index,
            distance_in=LANE_WIDTH_IN,
            start_conditions=start_conditions,
            restart=restart_mode,
        )
        for index, (runner, restart_mode) in enumerate(
            zip(runners, ("always", "whenNotActive", "never"), strict=True),
            start=1,
        )
    )
    return _build_auto_timing_xml(
        start_id=max_shape_id + 1,
        main_effect_pars=effects,
        build_entries=tuple(
            BuildEntry(shape_id=runner.shape_id, grp_id=index)
            for index, runner in enumerate(runners, start=1)
        ),
    )


def build_timing_oracle_deck(output_path: Path) -> Path:
    return materialize_scaffold_package("timing_oracle", output_path)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tmp/powerpoint-timing-oracle-deck.pptx"),
        help="Output PPTX path.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable debug logging.",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    output = build_timing_oracle_deck(args.output)
    logger.info("Built timing oracle deck: %s", output)
    logger.info("Slide 2: end=2s should stop early.")
    logger.info("Slide 3: STOP button should end only the lower runner.")
    logger.info("Slide 4: repeatDur=2.5s should cap the lower loop.")
    logger.info("Slide 5: only restart=always should restart during overlap.")
    logger.info("Slide 6: restart=never should ignore the 5s second begin.")


if __name__ == "__main__":
    main()
