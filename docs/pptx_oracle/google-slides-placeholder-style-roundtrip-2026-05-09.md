# Google Slides Placeholder Style Roundtrip - 2026-05-09

This note records a DocsServiceApp live probe into Google Slides placeholder
font carriers. The initial question was whether Google Slides import/export
generally loses placeholder font settings. It does not; the carrier matters.

## Generic Google-Authored Template Probe

A native Google Slides deck was created with the standard title slide layout.
The slide placeholders were styled through `SlidesApp`; the inherited layout
placeholders were also styled through the Slides advanced service where
possible. The deck was exported to PPTX, imported back into Google Slides, and
exported again.

Artifacts:

```text
source:   https://docs.google.com/presentation/d/1anFrxOJUWRprTuj1rzB_XNwLrrjUyzB-aGrM1YVJPG8/edit
imported: https://docs.google.com/presentation/d/1x9-n0YhWbCXxklXhy6TyLVsSDu52Zh3K0HOqJ3aOsBE/edit
```

Observed result:

| Carrier | Native Google export | After import/export | Interpretation |
|---|---|---|---|
| slide placeholder run fonts | `Courier New`, `Georgia` | `Courier New`, `Georgia` | preserved |
| layout placeholder default fonts | `Courier New`, `Georgia` | `Arial`, `Courier New`, `Georgia` | requested fonts preserved; Google may add Arial |

Conclusion: Google does not inherently discard Google-authored placeholder font
settings. Slide placeholder run-level style survives. Layout placeholder
defaults can survive, though Google may add normalized fallback/default entries.

## Generated PPTX / Figma Import Probe

DocsServiceApp generated a one-slide PPTX from a Figma-exported SVG. The
generated PPTX places non-text designer geometry on the slide layout and text on
placeholder-bearing slide shapes. The first implementation wrote placeholder
defaults as paragraph-level `a:pPr/a:defRPr`; Google imported the file, but the
exported layout placeholder typefaces normalized to `Arial`.

The carrier was then changed to the Google-style layout list-style path:

```xml
<a:lstStyle>
  <a:lvl1pPr>
    <a:defRPr sz="..." b="...">
      <a:solidFill>...</a:solidFill>
      <a:latin typeface="..."/>
    </a:defRPr>
  </a:lvl1pPr>
</a:lstStyle>
```

Smoke deck after the fix:

```text
https://docs.google.com/presentation/d/1RCtsBvTOub-ZdqZ4SM_DURDKmb1Qe_2UsJZpZ8L_2Co/edit
```

Observed result:

| Metric | Generated source PPTX | Google-exported PPTX |
|---|---|---|
| text/run typefaces | `Inter Tight`, `Space Mono`, `Syncopate` | `Arial`, `Inter Tight`, `Space Mono`, `Syncopate` |
| layout placeholder default typefaces | `Inter Tight`, `Space Mono`, `Syncopate` | `Arial`, `Inter Tight`, `Space Mono`, `Syncopate` |
| slide placeholder default typefaces | `Inter Tight`, `Space Mono`, `Syncopate` | none exported as placeholder defaults |
| slide placeholders | 4 | 4 |
| layout placeholders | 30 | 59 |
| layout images | 2 | 2 |

Conclusion: generated layout placeholder font families survive Google
Slides import/export when they are authored as list-style level defaults
(`a:lstStyle/a:lvl1pPr/a:defRPr`). Google still strips slide-placeholder
default style carriers and may add Arial to layout defaults. Direct run
styling on visible text remains the most reliable visual-fidelity carrier.

## Oracle / Authoring Rules

- Do not classify an added `Arial` placeholder default as loss when requested
  typefaces still survive in the exported placeholder defaults or text runs.
- For Google-targeted PPTX generation, write master/layout placeholder text
  defaults through `a:lstStyle/a:lvlNpPr/a:defRPr`, not only through paragraph
  `a:pPr/a:defRPr`.
- Keep explicit `a:rPr` styling on actual text runs when visible fidelity
  matters. This survives Google import/export more reliably than slide
  placeholder defaults.
- Treat slide placeholder default style stripping as converter normalization,
  not necessarily user-visible font loss, if the actual text run style survives.
