# Google Slides Animation Roundtrip - 2026-05-08

Live file-level Google Slides import/export probe for PresentationML
transitions and timing effects.

## Scope

This probe covers the current DocsServiceApp/OpenXML PowerPoint-style authoring
surface and the known animation-oracle vocabulary:

- 21 base `p:` slide transition elements.
- 19 Office 2010 `p14:` slide transition elements.
- 1 Office 2013 `p15:prstTrans` smoke case, `prst="fallOver"`.
- 56 entrance/exit `p:animEffect` filter cases: the 28 entrance/exit filters
  from `svg2ooxml/src/svg2ooxml/assets/animation_oracle/filter_vocabulary.xml`,
  excluding `image`.
- 1 `p:animEffect filter="image"` opacity-emphasis case.

The probe generated valid source PPTX decks, uploaded each deck through Drive
as a Google Slides presentation, exported each native presentation back to
PPTX, and inspected the resulting slide XML. It is a file-level export oracle;
it does not prove browser slideshow playback behavior inside Google Slides.

For the opposite direction, where the source is a native Google Slides deck
exported to PPTX, see
`docs/pptx_oracle/google-slides-native-export-oracle-2026-05-08.md`. That
native export oracle is the better target when authoring PPTX intended for a
Google Slides workflow.

Local artifacts from the run:

```text
/private/tmp/google-slides-animation-roundtrip-2026-05-08/
  base-transitions.pptx
  base-transitions.google.pptx
  p14-transitions.pptx
  p14-transitions.google.pptx
  p15-transition.pptx
  p15-transition.google.pptx
  effects.pptx
  effects.google.pptx
  roundtrip-report.json
```

Validation command:

```bash
/Users/ynse/projects/openxml-audit/.venv/bin/openxml-audit --recursive /private/tmp/google-slides-animation-roundtrip-2026-05-08
```

Result: all source and exported PPTX packages reported `Errors: 0`.

## Summary

| Surface | Cases | Preserved | Rewritten | Dropped |
|---|---:|---:|---:|---:|
| Base `p:` transitions | 21 | 2 | 19 | 0 |
| `p14:` transitions | 19 | 0 | 14 | 5 |
| `p15:prstTrans` | 1 | 0 | 0 | 1 |
| Entrance/exit filter effects | 56 | 2 | 54 | 0 |
| `filter="image"` opacity emphasis | 1 | 0 | 0 | 1 |

Practical result for PowerPoint-style imported PPTX: Google Slides preserves
only the small carrier subset:
`p:fade`, `p:push`, and `p:animEffect` fade in/out. It preserves timing
milliseconds for the surviving fade animations, but it normalizes transition
and animation XML heavily.

## Slide Transitions

Every source transition used `spd="fast"`. Google dropped the `spd` attribute
in every exported transition.

Base `p:` transitions:

| Source | Exported | Verdict |
|---|---|---|
| `p:blinds` | `p:fade` | rewritten |
| `p:checker` | `p:fade` | rewritten |
| `p:circle` | `p:fade` | rewritten |
| `p:dissolve` | `p:fade` | rewritten |
| `p:comb` | `p:fade` | rewritten |
| `p:cover` | `p:push` | rewritten |
| `p:cut` | `p:fade` | rewritten |
| `p:diamond` | `p:fade` | rewritten |
| `p:fade` | `p:fade` | preserved |
| `p:newsflash` | `p:fade` | rewritten |
| `p:plus` | `p:fade` | rewritten |
| `p:pull` | `p:push` | rewritten |
| `p:push` | `p:push` | preserved |
| `p:random` | `p:fade` | rewritten |
| `p:randomBar` | `p:fade` | rewritten |
| `p:split` | `p:fade` | rewritten |
| `p:strips` | `p:fade` | rewritten |
| `p:wedge` | `p:fade` | rewritten |
| `p:wheel` | `p:fade` | rewritten |
| `p:wipe` | `p:fade` | rewritten |
| `p:zoom` | `p:fade` | rewritten |

Most non-fade base transitions exported as `<p:fade thruBlk="1"/>`. `cover`
and `pull` exported as `p:push`.

Office 2010 `p14:` transitions:

| Source | Exported | Verdict |
|---|---|---|
| `p14:flash` | `p:fade` | rewritten |
| `p14:vortex` | `p:fade` | rewritten |
| `p14:switch` | dropped | dropped |
| `p14:flip` | dropped | dropped |
| `p14:ripple` | `p:fade` | rewritten |
| `p14:glitter` | `p:fade` | rewritten |
| `p14:honeycomb` | `p:fade` | rewritten |
| `p14:prism` | dropped | dropped |
| `p14:doors` | `p:fade` | rewritten |
| `p14:window` | `p:fade` | rewritten |
| `p14:shred` | `p:fade` | rewritten |
| `p14:ferris` | `p:fade` | rewritten |
| `p14:flythrough` | `p:fade` | rewritten |
| `p14:warp` | `p:fade` | rewritten |
| `p14:gallery` | dropped | dropped |
| `p14:conveyor` | dropped | dropped |
| `p14:pan` | `p:fade` | rewritten |
| `p14:reveal` | `p:fade` | rewritten |
| `p14:wheelReverse` | `p:fade` | rewritten |

No `p14:` transition survived as `p14:`. The `p15:prstTrans` smoke case
roundtripped with no transition.

## Timing Effects

The source effects used:

- `nodeType="clickEffect"`
- outer start condition `delay="250"`
- `p:animEffect` duration `dur="1500"`
- entrance and exit directions through `transition="in"` or `transition="out"`
- visibility `p:set` primers for entrance/exit filter effects

Google export behavior:

- `fade` entrance and exit stayed as `filter="fade"` with the same direction.
- Every other entrance/exit filter exported as `filter="fade"` with the same
  `transition="in"` or `transition="out"`.
- `dur="1500"` was preserved on exported `p:animEffect` nodes.
- The start condition delay `250` was preserved, but Google rewrote the timing
  container from `nodeType="clickEffect"` to `nodeType="withEffect"`.
- Google normalized preset UI metadata to `presetID="10"` / `presetSubtype="0"`
  for both entrance and exit fade exports.
- Visibility primers survived. Exit visibility-set delay changed from `1499`
  to `1500`, which is equivalent to hiding at the end of the 1500 ms effect.
- Shape ids were remapped and timing targets were updated.
- `p:animEffect filter="image" prLst="opacity: 0.3"` plus `style.opacity`
  primer was dropped.

The non-fade filter values that collapsed to fade were:

```text
dissolve
wipe(down) wipe(up) wipe(left) wipe(right)
wedge
wheel(1) wheel(2) wheel(3) wheel(4) wheel(8)
circle(in) circle(out)
strips(downLeft) strips(downRight) strips(upLeft) strips(upRight)
blinds(horizontal) blinds(vertical)
checkerboard(across) checkerboard(down)
barn(inVertical) barn(inHorizontal) barn(outVertical) barn(outHorizontal)
randombar(horizontal) randombar(vertical)
```

Each value was tested in both entrance and exit direction.

## Authoring Rule

For PowerPoint-style PPTX import/export compatibility:

- Use `p:fade` or `p:push` for slide transitions. Treat all other transition
  kinds, all `p14:` transition kinds, all `p15:` preset transitions, transition
  speed, and original shape ids as non-stable.
- Use `p:animEffect transition="in|out" filter="fade"` for animations that
  must survive as the same effect kind.
- Do not depend on non-fade `animEffect` filters surviving; they roundtrip as
  fade.
- Do not use `filter="image"` opacity emphasis for Google Slides
  compatibility; it was removed.
- Continue serializing timing values as OOXML millisecond integers. The Google
  exporter preserved `dur="1500"` and the start delay in this sweep.

For Google-targeted authoring, prefer the Google-native export shapes recorded
in `google-slides-native-export-oracle-2026-05-08.md`.
