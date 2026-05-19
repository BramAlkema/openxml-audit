# Google Docs Linked Text Box Evidence

Date: 2026-05-18

This note captures a narrow Google Docs import/export probe for OOXML linked
text-box stories. The feature exists in the WordprocessingShape schema as
`wps:linkedTxbx`, but Google Docs does not preserve that link marker when a DOCX
is imported as a Google Doc and exported back to DOCX.

## Source Shape

The generated DOCX contained one inline `wpg:wgp` drawing with two
`wps:wsp` text-box shapes:

- first shape: `wps:txbx id="1"` with visible `w:txbxContent`;
- second shape: `wps:linkedTxbx id="1" seq="1"`.

The intent is the OOXML linked text-box story model: text starts in the first
box and the second box participates in the same story by reference.

## Google Docs Roundtrip Result

The probe imported the DOCX through Google Drive conversion, exported the
resulting Google Doc back to DOCX, inspected `word/document.xml`, and deleted
the temporary Drive file.

Observed result:

- `wpg:wgp` survived;
- `wps:wsp` survived;
- visible `w:txbxContent` text survived;
- `wps:linkedTxbx` did not survive;
- exported DOCX contained zero `wps:linkedTxbx` elements.

## Compatibility Conclusion

Classify `wps:linkedTxbx` as **schema-valid but Google-stripped** for the Google
Docs roundtrip oracle until contradicted by a richer fixture. Do not use linked
text-box stories as a Google Docs carrier for editable text flow.

For Google-backed transformations:

- preserve simple standalone `wps:txbx` content when the target is DOCX;
- lower simple text boxes only when direct text/layout mapping is sufficient;
- render or flatten linked text-box stories when exact flow semantics matter.

This is separate from the high-resolution drawing-canvas carrier: simple WPG/WPS
geometry can survive Google Docs roundtrips, but the linked text-box story marker
does not.
