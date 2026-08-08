# Spec 038 — EuroOffice and Google Docs DOCX differential oracle

## Status

Implemented on 2026-08-08. The EuroOffice editor transport has live evidence;
the Google Docs DOCX route is implemented and covered with a fake Drive client,
but a paired live baseline still requires an explicit impersonation subject and
owned staging-folder ID.

## Question

Can the same openly licensed DOCX survive a genuine EuroOffice editor save and
a Google Docs import/export, and which document semantics form the portable
intersection between those targets?

This is not answered by byte equality. Both editors rewrite OPC packages. It
is also not answered by treating Google as the normative oracle. The original
DOCX is the shared base; EuroOffice and Google are independent targets.

## Evidence layers

1. **Transport** — did the editor really open and persist the document?
2. **Package diff** — which canonical XML parts changed, appeared, or vanished?
3. **Document semantics** — did content order, headings, lists, tables,
   sections, headers/footers, fields, styles, theme, metadata, and security
   surfaces survive?
4. **Schema/security** — is the editor-produced package still accepted by the
   strict validator?
5. **Differential matrix** — for every semantic family, which targets preserve
   the source and what is preserved by all targets?

These layers remain separate. A callback save proves transport, not survival.

## EuroOffice transport

`tools/oracle/eurooffice_roundtrip.py` drives the Apache-licensed example app
bundled with the deployed EuroOffice Document Server. A pinned Playwright
browser shares the server container's network namespace because the example
app generates loopback callback/download URLs.

One pass would either leave no proof of a save or alter source content. The
transport therefore performs two dirty sessions:

1. upload under a UUID filename;
2. open in the editor, insert a UUID marker, require EuroOffice's own
   document-state event to report dirty, disconnect, and poll until the stored
   package hash changes;
3. reopen, remove exactly that marker, again require dirty state, disconnect,
   and poll for a second package hash;
4. download the final DOCX and delete the remote example-app file.

The marker never appears in the intended final semantics. The final output is
still a two-save stress result, which is conservative for survival claims.

### Requirements and boundary

The live transport is developer/lab infrastructure. It requires SSH and rsync
on the runner, Docker on the target host, and the bundled EuroOffice example
app enabled inside the configured Document Server container. The example app
itself warns that it is for testing rather than production use; do not expose
it as an ordinary document store. The browser image and Playwright package are
pinned to `v1.62.0`; the first run needs registry/npm access unless those
artifacts are already cached. Host, container, and browser image are explicit
CLI options.

## Google Docs transport

Spec 031's Drive client already carried DOCX and native Google Docs MIME types.
Its orchestrator now detects DOCX, uploads it to the caller-supplied staging
folder, converts it to a native Google document, exports DOCX, computes the
same semantic snapshot, and deletes both Drive-side files in `finally`.

The paired CLI validates the Google credential path, impersonation subject,
and folder ID before starting EuroOffice. It does not guess a Workspace user
or upload to Drive root.

## DOCX semantic snapshot

`openxml_audit.docx.semantic_snapshot` extracts eleven independently hashed
feature families:

- visible block content and order;
- resolved heading hierarchy;
- list level and numbering definitions;
- table cells, header rows, style, and look flags;
- page/section geometry, columns, first-page policy, and references;
- header/footer content;
- field instructions;
- paragraph, character, and table style semantics;
- theme colors and major/minor fonts;
- stable core/custom metadata;
- macros, embedded objects, and external relationships.

ZIP timestamps, XML namespace prefixes, relationship IDs, and volatile
modified/creator metadata are deliberately not semantic gates. Canonical XML
part diffs remain available beside the semantic result.

## Commands

```bash
# EuroOffice only
python -m openxml_audit.oracle eurooffice input.docx --keep-artifacts

# Google Docs only (PPTX and DOCX are accepted)
python -m openxml_audit.oracle gsuite input.docx --keep-artifacts

# Compare existing outputs without any live mutation
python -m openxml_audit.oracle docx-diff input.docx \
  --target eurooffice=eurooffice.docx \
  --target google_docs=google.docx

# Run both and build the same matrix
python -m openxml_audit.oracle docx-paired input.docx --output report.json
```

## First live EuroOffice finding

The open-source `acme-us.docx` corpus fixture completed both dirty editor
passes. Its package hash changed from
`900bcb1349b385871a2fb7cfac6a03003ffb998675bfc63acbff2ef4d3c7042a` to
`749896a0ad07484e4995b457112b038c6454d6d176f267f61f8367668d46dd91`.
Visible content and heading hierarchy survived. Section policy, style
semantics, and theme font fallbacks changed. A clean strict-validation input
became an output with at least 251 reported schema/security findings (the live
run's configured collection cap was reached). This is evidence
of a real editor save and a failed preservation gate, not evidence that every
finding necessarily predicts EuroOffice load failure.

## Acceptance criteria

- EuroOffice evidence requires two opened, dirty sessions and two distinct
  callback-produced hashes.
- Google DOCX uses native Google Docs conversion and cleans up Drive files.
- Both engines use the same DOCX semantic snapshot.
- Differential output reports per-target survival and the portable
  intersection without elevating either editor to source-of-truth status.
- Corpus entries have explicit license, provenance, and hashes.
- Live credentials and remote documents are never committed.
