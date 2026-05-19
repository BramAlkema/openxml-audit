# Google Editor Clipboard Text Box Evidence

Date: 2026-05-16

This fixture pair captures a Google editor copy/paste JSON route, not a generic
Word-to-PowerPoint OOXML conversion. It is useful as evidence for Google's
clipboard lowering surface: a Google-authored drawing/text-box object can appear
as a WordprocessingML text box in DOCX and as a native PresentationML text box in
PPTX.

## Files

- `data/corpus/gsuite_native_export/docx/google-docs-clipboard-textbox-2026-05-16.docx`
  - Source: `/Users/ynse/Downloads/test2.docx`
  - SHA-256: `742591a6a4628d9ee8da86662f37e54e2391c3e6183d8cfb7b9c96e981a339d6`
  - Package: 10 parts
  - Key shape: one standalone `wps:wsp` text box in `mc:AlternateContent`
  - Fallback: one 1 x 1 PNG at `word/media/image1.png`
  - Current validator result: 6 errors, all on the Google-authored text-box
    branch (`wps:wsp` child order / missing required elements in the schema
    model, plus decimal `w:spacing/@w:line` values)
- `data/corpus/gsuite_native_export/pptx/google-slides-clipboard-textbox-2026-05-16.pptx`
  - Source: `/Users/ynse/Downloads/test1.pptx`
  - SHA-256: `b704b3cd76c524f82615383bff2daa15ee6fa271a91a93bb308c29da0901a372`
  - Package: 46 parts
  - Key shape: one native `p:sp` text box with `p:txBody`
  - Media: no `ppt/media/*` parts
  - Current validator result: 8 errors, all invalid embedded font payload
    reports for Google-exported `.fntdata` parts

## Observed Mapping Surface

The DOCX side carries visible text in `w:txbxContent` paragraphs and runs. The
PPTX side carries the same visible text in DrawingML paragraphs and runs:
`p:sp/p:txBody/a:p/a:r/a:t`.

Useful direct mappings visible in this fixture:

- `wps:spPr/a:xfrm` to `p:spPr/a:xfrm`
- `wps:bodyPr` to `a:bodyPr`
- `w:spacing` to `a:spcBef`, `a:spcAft`, and `a:lnSpc`
- `w:jc` to `a:pPr/@algn`
- `w:rFonts`, `w:b`, `w:i`, `w:color`, and `w:sz` to `a:rPr`

This is enough evidence for a narrow `lowerSimpleTextBox()` path for
Google-authored clipboard/export surfaces.

## Caution

Do not generalize this fixture to arbitrary DOCX text boxes. The PPTX contains
PowerPoint list metadata such as `a:buAutoNum`, while the DOCX text-box branch
mainly exposes direct runs and indentation rather than a clean `w:numPr`
numbering contract. A mapper using this evidence should emit a loss report for
numbering, tabs, fields, inherited styles, complex text direction, and anything
not backed by explicit direct formatting.

For exact Word text-box semantics, preserve the original DOCX `wps:wsp` /
`w:txbxContent` island or render it instead of lowering it to PresentationML.
