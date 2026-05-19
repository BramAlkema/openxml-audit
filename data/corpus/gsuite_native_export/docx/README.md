# Google Docs Native Export Corpus

Small DOCX files exported from Google Docs or produced through Google editor
copy/paste surfaces. These fixtures document how Google serializes its own
document/drawing features back to OOXML.

## Fixtures

- `google-docs-clipboard-textbox-2026-05-16.docx`
  - Source: user-provided `/Users/ynse/Downloads/test2.docx`.
  - SHA-256: `742591a6a4628d9ee8da86662f37e54e2391c3e6183d8cfb7b9c96e981a339d6`.
  - Contents: Google editor copy/paste JSON text-box surface exported as a
    standalone `wps:wsp` text box with a 1 x 1 PNG fallback.
  - Validation: current `openxml-audit` reports 6 schema errors in the
    Google-authored text-box branch.
  - Analysis: `docs/gsuite_oracle/google-editor-clipboard-textbox-2026-05-16.md`.
