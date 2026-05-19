# Google Slides Native Export Oracle - 2026-05-08

Analysis of a user-provided PPTX exported directly from Google Slides:

```text
data/corpus/gsuite_native_export/pptx/google-slides-effects-transitions-2026-05-08.pptx
```

The original file was `/Users/ynse/Downloads/test3.pptx`; the copied corpus
fixture has SHA-256
`b96285625afe52d187168e7e15490a1889b2542fb4d8118d8f0456174b06fb86`.

This is a different oracle from the PPTX import/export sweep in
`google-slides-animation-roundtrip-2026-05-08.md`: that sweep started with
PowerPoint-style PresentationML and observed how Google rewrote it. This file
starts from Google's own native animation/transition model and shows how Google
serializes that model to PresentationML.

Validation:

```bash
/Users/ynse/projects/openxml-audit/.venv/bin/openxml-audit \
  data/corpus/gsuite_native_export/pptx/google-slides-effects-transitions-2026-05-08.pptx
```

Result: `Errors: 0`.

## Native Transition Exports

The deck has seven slides. Four slides contain transition XML; three export
with no `<p:transition>` element.

| Slide | Exported transition |
|---:|---|
| 1 | `<p:transition spd="med"><p:fade/></p:transition>` |
| 2 | `<p:transition spd="med"><p:fade thruBlk="1"/></p:transition>` |
| 3 | `<p:transition spd="med"><p:push/></p:transition>` |
| 4 | `<p:transition spd="med"><p:push dir="r"/></p:transition>` |
| 5 | none |
| 6 | none |
| 7 | none |

The source deck does not carry human-readable transition labels, so map slide
numbers back to the Google UI only when the source deck ordering is known.
Structurally, the Google-native export surface is limited to two fade forms,
two push forms, and absence of transition.

## Native Effect Exports

All animation timing is on slide 1. Google exports 12 effect containers:

| Preset class | Preset ID | Subtype | XML behavior signature | Target attribute/filter |
|---|---:|---:|---|---|
| `entr` | 1 | 0 | `set` | `style.visibility -> visible` |
| `exit` | 1 | 0 | `set` | `style.visibility -> hidden` |
| `entr` | 10 | 0 | `set + animEffect` | `filter="fade" transition="in"` |
| `exit` | 10 | 0 | `animEffect + set` | `filter="fade" transition="out"` |
| `entr` | 2 | 8 | `set + anim` | `ppt_x: #ppt_x-1 -> #ppt_x` |
| `entr` | 2 | 2 | `set + anim` | `ppt_x: #ppt_x+1 -> #ppt_x` |
| `entr` | 2 | 4 | `set + anim` | `ppt_y: #ppt_y+1 -> #ppt_y` |
| `entr` | 2 | 1 | `set + anim` | `ppt_y: #ppt_y-1 -> #ppt_y` |
| `exit` | 2 | 1 | `anim + set` | `ppt_y: #ppt_y -> #ppt_y-1` |
| `entr` | 23 | 16 | `set + anim + anim` | `ppt_w/ppt_h: 0 -> current` |
| `exit` | 23 | 32 | `anim + anim + set` | `ppt_w/ppt_h: current -> 0` |
| `emph` | 8 | 0 | `animRot` | `r`, `by="-21600000"` |

Common timing shape:

- `nodeType="clickEffect"`
- container start condition `delay="0"`
- behavior duration `dur="1000"` for fade/motion/size/rotation behaviors
- visibility primers use `p:set` on `style.visibility`

Behavior primitive counts:

| Primitive | Count |
|---|---:|
| `p:set` | 11 |
| `p:anim` | 9 |
| `p:animEffect` | 2 |
| `p:animRot` | 1 |

Attribute vocabulary observed:

| Behavior | Attribute | Count |
|---|---|---:|
| `set` | `style.visibility` | 11 |
| `anim` | `ppt_y` | 3 |
| `anim` | `ppt_x` | 2 |
| `anim` | `ppt_w` | 2 |
| `anim` | `ppt_h` | 2 |
| `animRot` | `r` | 1 |

## Reimport Check

The Google-native export was imported back into Google Slides and exported
again to:

```text
/private/tmp/google-slides-native-oracle-test3-roundtrip.pptx
```

The second export validated with `Errors: 0`. The semantic timing vocabulary
was stable:

- same seven-slide transition surface
- same 12 effect containers
- same preset class/id/subtype signatures
- same behavior primitive counts
- same `animEffect` filters and directions
- same `ppt_x`, `ppt_y`, `ppt_w`, `ppt_h`, `r`, and `style.visibility`
  attribute vocabulary

The only normalized timing delta in the normalized timing diff was the
`exit`/`presetID="1"` visibility-set child delay changing from `0` to `1`.
The package still had broad XML/defaults churn across themes, layouts, notes,
and slides, as expected for a Google import/export cycle.

## Elaboration Probe

A follow-up generated PPTX probe tested whether more elaborate `p:animRot` and
`p:anim` forms survive a Google Slides import/export. Artifacts:

```text
/private/tmp/google-slides-animrot-anim-probe-2026-05-08/
  animrot-anim-elaboration.pptx
  animrot-anim-elaboration.google.pptx
  report.json
```

Both source and exported PPTX validated with `Errors: 0`.

Result: the effect families survived, but Google canonicalized the details to
its native preset shapes.

| Probe | Source | Google export |
|---|---|---|
| clockwise 360 rotation | `animRot by="21600000"` | `animRot by="-21600000"` |
| counter-clockwise 180 rotation | `animRot by="-10800000"` | `animRot by="-21600000"` |
| clockwise 720 rotation | `animRot by="43200000"` | `animRot by="-21600000"` |
| rotation from/to | `animRot from="0" to="21600000"` | `animRot by="-21600000"` |
| far X motion | `#ppt_x-2 -> #ppt_x` | `#ppt_x-1 -> #ppt_x` |
| far Y motion | `#ppt_y+2 -> #ppt_y` | `#ppt_y+1 -> #ppt_y` |
| 3-keyframe X motion | `#ppt_x-1 -> #ppt_x+1 -> #ppt_x` | `#ppt_x-1 -> #ppt_x` |
| diagonal X+Y motion | `ppt_x` and `ppt_y` siblings | only `ppt_x` survived |
| half-size zoom | `#ppt_w/2`, `#ppt_h/2 -> current` | `0 -> #ppt_w`, `0 -> #ppt_h` |
| overshoot zoom | `0 -> #ppt_w*1.2 -> #ppt_w` | `0 -> #ppt_w` |

Google also normalized these elaborated effect containers from
`nodeType="clickEffect"` to `nodeType="withEffect"`.

Practical interpretation: richer `p:anim` / `p:animRot` XML is accepted and
does not break import/export, but it should be treated as a request for one of
Google's coarse native presets, not as a precise custom animation curve.

## Shape Effect Import/Export Probe

A DocsServiceApp DrawingML designer smoke on 2026-05-09 tested native
DrawingML shape effects through the same Google Slides import/export path. The
probe imported a one-slide PPTX with nine authored shapes, then exported the
native Google Slides deck back to PPTX.

Live deck:

```text
https://docs.google.com/presentation/d/1Lyffiz2lrWKnB94iEzdwM686DKkcnQsfTkGnaT8SnmM/edit
```

Element counts:

| Element | Source PPTX | Google-exported PPTX | Verdict |
|---|---:|---:|---|
| `a:gradFill` | 9 | 8 | mostly preserved |
| `a:outerShdw` | 6 | 6 | preserved |
| `a:reflection` | 1 | 1 | preserved |
| `a:glow` | 6 | 0 | stripped |
| `a:softEdge` | 1 | 0 | stripped |

Practical interpretation: Google Slides accepts a PPTX containing glow and
soft-edge DrawingML, but those effects are not stable export carriers. Treat
`a:glow` and `a:softEdge` as lossy for Google-targeted PPTX generation. Use
native gradients, outer shadows, and reflection when the result must survive
Google Slides import/export.

## Authoring Implication

For Google Slides compatibility there are two different surfaces:

1. **PowerPoint-style PPTX imported into Google.** Non-fade
   `p:animEffect filter="..."` values collapse to fade, and p14/p15
   transitions are not stable.
2. **Google-native PPTX shape.** Google exports and reimports a stable small
   animation vocabulary based on `set`, `animEffect fade`, `anim` over
   `ppt_x/ppt_y/ppt_w/ppt_h`, and `animRot` over `r`.

TokenMoulds/DocsServiceApp should use the Google-native shape, not the broad
PowerPoint filter vocabulary, when the target is a Google Slides import/export
workflow. Keep the parameters canonical: one-unit fly-style `ppt_x`/`ppt_y`,
zero-to-current zoom-style `ppt_w`/`ppt_h`, and one-spin `animRot`.
