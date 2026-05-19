# Google Slides Native Export Corpus

Small PPTX files exported directly from native Google Slides presentations.
These fixtures document how Google serializes its own presentation features
back to OOXML.

## Fixtures

- `google-slides-effects-transitions-2026-05-08.pptx`
  - Source: user-provided `/Users/ynse/Downloads/test3.pptx`.
  - SHA-256: `b96285625afe52d187168e7e15490a1889b2542fb4d8118d8f0456174b06fb86`.
  - Contents: Google Slides export oracle for the native transition/effect
    surface available in the source Google Slides deck.
  - Validation: `openxml-audit` reports `Errors: 0`.
  - Analysis: `docs/pptx_oracle/google-slides-native-export-oracle-2026-05-08.md`.

- `google-slides-clipboard-textbox-2026-05-16.pptx`
  - Source: user-provided `/Users/ynse/Downloads/test1.pptx`.
  - SHA-256: `b704b3cd76c524f82615383bff2daa15ee6fa271a91a93bb308c29da0901a372`.
  - Contents: Google editor copy/paste JSON text-box surface exported as a
    native `p:sp` text box with no `ppt/media/*` parts.
  - Validation: current `openxml-audit` reports 8 invalid embedded font
    payload errors for Google-exported `.fntdata` parts.
  - Analysis: `docs/gsuite_oracle/google-editor-clipboard-textbox-2026-05-16.md`.
