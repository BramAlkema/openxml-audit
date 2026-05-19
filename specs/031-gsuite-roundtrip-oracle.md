# Spec: Google Workspace Roundtrip Oracle

## Status

Proposed (May 3, 2026). Phase 1 candidate for the next 0.7.x release —
adds a fifth engine to the oracle dispatcher
(`src/openxml_audit/oracle/__main__.py`) alongside Word, ODF,
PowerPoint, and Excel.

## Problem

The validator's mission is roundtrip survival in the *target app*.
Today the oracle layer covers four targets:

- Word for Mac via osascript (Spec 011)
- LibreOffice / soffice headless (Spec 019)
- PowerPoint for Mac via osascript (Spec 020)
- Excel for Mac via osascript (Spec 021)

Google Workspace (Slides, Docs, Sheets) is the third major consumer
of OOXML in the real world — and the only target that imports OOXML
into a *different native format* (Google's proprietary IR) and
re-exports. The conversion is lossy *by design*: features that don't
map to Google's IR are dropped, simplified, or transformed.

Without a GSuite oracle, the evidence ladder is silent on the
most user-visible loss surface for non-MS-Office workflows.

## Why This Matters

- **Evidence-ladder coverage.** Project mission: scope is "will this
  survive a roundtrip?" — GSuite is the third leg of the stool
  alongside Office and ODF/LibreOffice.
- **Canonical reference documents.** A canonical .pptx that survives
  PowerPoint *and* LibreOffice *and* GSuite is the strongest possible
  "this works everywhere" claim — and the project's stated ultimate
  deliverable.
- **Loss as signal.** GSuite drops features by design. A public
  catalogue of "GSuite doesn't roundtrip these OOXML constructs"
  is itself a deliverable that doesn't exist anywhere today.

## Normative References

- `tools/oracle/pptx_repair_oracle.py` — sibling oracle, target
  observation shape.
- `tools/oracle/odf_repair_oracle.py` — sibling oracle, headless
  conversion pattern (`soffice --convert-to`).
- `src/openxml_audit/oracle/__main__.py` — dispatcher this spec adds
  an engine to.
- `src/openxml_audit/pptx/lab.py` — `compare_pptx_packages` per-part
  diff infrastructure this spec consumes.
- `../tokenmoulds/src/tokenmoulds/mcp/tools/google_apps_script.py`
  — sibling project's clasp/manifest/scopes scaffolding. Reference
  only; this spec does **not** use Apps Script (auth model differs —
  see Approach).
- Google Drive API v3: `files.create`, `files.copy` with conversion,
  `files.export`, `files.delete`.

## Approach

### Auth model — service account with domain-wide delegation

One auth principal: the project owner. A service account in a
project-owned GCP project, one JSON key at
`~/.config/openxml-audit/google_service_account.json` (override via
`GSUITE_ORACLE_CREDS`). The service account impersonates a real
Workspace user via domain-wide delegation; the impersonation subject
defaults to the value of `GSUITE_ORACLE_SUBJECT`.

Domain-wide delegation is **mandatory**, not optional, for two reasons
discovered empirically during smoke-testing:

1. **Service accounts have zero storage quota** since Google's 2024
   policy change. `files.create` against an SA's own Drive returns
   `storageQuotaExceeded`. Even a regular user-owned folder shared
   with the SA cannot accept SA-authored uploads — the upload is
   attributed to the SA, which has no quota anywhere.
2. **Shared Drives are a Workspace tier feature** (Business Standard+).
   They work, but require an extra paid tier. Domain-wide delegation
   works on every Workspace tier including Business Starter.

Under domain-wide delegation, the SA acts *as* the impersonated user,
files are attributed to that user's Drive, and uploads consume that
user's quota — the standard pattern for headless server-side workflows.

Setup is one-time:

1. In GCP Console → IAM & Admin → Service Accounts → `oracle-roundtrip`
   → "Show domain-wide delegation," copy the OAuth client ID.
2. In Workspace Admin Console → Security → API controls →
   Domain-wide Delegation → Add new, paste the client ID with scope
   `https://www.googleapis.com/auth/drive`.
3. Drop the service account JSON key at
   `~/.config/openxml-audit/google_service_account.json` (`chmod 600`).
4. Set `GSUITE_ORACLE_SUBJECT` to the user the SA should impersonate
   (e.g., the project owner's Workspace email).
5. Create a regular Drive folder owned by that user to hold in-flight
   oracle uploads; set `GSUITE_ORACLE_FOLDER_ID` to its folder ID.

No end-user OAuth. No Apps Script Web App. The Apps Script route from
`../tokenmoulds` solves a different problem (per-customer auth for
template emission); the oracle's single-principal world makes domain-
wide delegation simpler — no deploy, no 30-second response cap, no
shared secret to manage.

### Phase 1 — PPTX only

1. **`src/openxml_audit/gsuite/client.py`** — thin Drive API wrapper:
   - `upload(path) -> file_id` — upload as raw OOXML.
   - `convert_to_native(file_id, target_mime) -> file_id` —
     `files.copy` with the Google native mime type. This is the
     "import as Slides" step.
   - `export_to_ooxml(file_id, ooxml_mime) -> bytes` — `files.export`
     back to OOXML.
   - `delete(file_id)` — best-effort cleanup.

2. **`tools/oracle/gsuite_roundtrip.py`** — orchestrator. Per file:
   - Stage original locally under `~/Documents/.gsuite_oracle_runs/<id>/`
     (mirrors sibling oracles' staging convention).
   - Upload → convert-to-native → export-to-ooxml → write download
     next to the original.
   - Hand original + post-GSuite copy to `compare_pptx_packages` —
     same per-part diff infrastructure as the PowerPoint oracle.
   - Roll up into a `RoundtripObservation` extended with a `loss`
     field of type `set[LossClass]` (see below).
   - Cleanup the three Drive files (best-effort; record failures).

3. **`LossClass` taxonomy** — Phase 1 buckets, rule-based over the
   per-part diff. Multiple classes may fire per file. Naming policy:
   stay descriptive. `*_part_changed` / `*_part_removed` say what
   we *observed in the diff* without claiming semantic loss. The
   `*_loss` names are *reserved* for future buckets that verify
   actual loss — e.g., `font_loss` would only fire if a font is
   removed AND was referenced by a text run; `style_loss` would
   only fire if `tableStyles.xml` is removed AND the source has
   `<a:tbl>` shapes using it. Until that verification logic exists,
   file-level signals get descriptive names so the oracle doesn't
   over-claim. (Example over-claim avoided: Presentation1.pptx has
   no tables, so Google removing the unreferenced `tableStyles.xml`
   isn't loss — it's pruning.)

   File-level descriptive buckets:
   - `theme_part_changed` — `ppt/theme/*.xml` changed or removed.
   - `master_part_changed` — slide masters/layouts changed or
     removed.
   - `style_part_removed` — `ppt/tableStyles.xml` (or future
     style-set parts) removed.
   - `font_part_removed` — embedded fonts dropped
     (`ppt/fonts/*` removed).
   - `slide_part_changed` — slide XML changed (often just
     re-formatting / namespace-bloat / shape-id renumbering).
   - `metadata_churn` — `docProps/*.xml` or `ppt/viewProps.xml`
     differences (often non-deterministic; isolated so it doesn't
     dominate other classes).
   - `media_re_encoded` — `ppt/media/*` bytes changed (genuine
     signal: image bytes did differ).
   - `structural_normalization` — GSuite *added* parts that weren't
     in the source (e.g., `ppt/notesMasters/*`, `ppt/notesSlides/*`,
     extra `ppt/theme/themeN.xml` variants). Lossless additions but
     still roundtrip artifacts.
   - `defaults_inlined` — content-aware signal. Source slides had
     empty inheriting elements (`<p:spPr/>`, `<a:bodyPr/>`,
     `<a:pPr/>`, `<a:cNvSpPr/>`, `<a:xfrm/>`) that the GSuite export
     replaced with explicit, fully-resolved values. **File-level
     observation only:** we can see that GSuite's exporter wrote
     resolved values, but cannot tell whether the semantic binding
     to the layout/master is broken inside Google's IR. The Slides
     app might still track inheritance internally and only inline-
     resolve on export. Verifying that requires a behavioral oracle
     (see "Out of Scope" below).

   Reserved (verified-loss) buckets — not currently fired:
   - `content_changed` — slide text or structural content differs
     beyond mechanical normalization. Lands once content-comparison
     logic exists.

   Catch-all:
   - `unmapped` — non-empty diff didn't fit any bucket; raw report
     attached for inspection.

   **Two classifier layers.** `classify_loss(...)` runs cheap
   list-based rules over the diff's part-name set. `classify_xml_loss(
   base_path, head_path, ...)` opens the packages and reads slide
   XML bytes for content-aware rules (`defaults_inlined` is the
   only Phase 1 rule). `observe()` unions the two. Tests can
   exercise either layer in isolation.

   Empirical baseline from a 32 KB single-slide Presentation1.pptx
   roundtrip: 4 parts removed, 5 parts added, 21 parts changed,
   output 4.5× larger. Classification: `metadata_churn` (docProps
   removed) + `style_part_removed` (unused tableStyles dropped) +
   `theme_part_changed` (theme1 rewritten) + `master_part_changed`
   (master+layouts re-emitted) + `slide_part_changed` (slide1
   rewritten) + `defaults_inlined` (empty placeholders expanded) +
   `structural_normalization` (notes master/slide added).

4. **Dispatcher wiring** — add `gsuite` (alias `google`) to
   `_DISPATCH` in `src/openxml_audit/oracle/__main__.py` so
   `python -m openxml_audit.oracle gsuite FILES...` routes correctly.

5. **Tests** — `tests/test_gsuite_roundtrip_oracle.py`:
   - Always-on: `LossClass` classifier against fixed diff fixtures,
     observation shape, dispatcher routing, CLI argument parsing,
     mocked Drive client.
   - GSuite-required smoke tests skip cleanly when `GSUITE_ORACLE_CREDS`
     is unset (CI without GCP creds passes).

### Phase 2 — DOCX, XLSX

Parameterize the orchestrator by format. Mime mapping:
- docx ↔ `application/vnd.google-apps.document`
- xlsx ↔ `application/vnd.google-apps.spreadsheet`

Per-format `LossClass` extensions: `formula_loss` (xlsx),
`style_loss` (docx), etc. Reuses each format's existing
`compare_*_packages` differ.

Empirical carrier matrix from the DOCX/XLSX/PPTX probes:

- **Google Docs -> DOCX** does not reliably export `customXml/*`
  sidecars (including `customXml/tokenmoulds-intent.json`),
  `docProps/core.xml`, `docProps/app.xml`, `word/webSettings.xml`,
  `word/stylesWithEffects.xml`, original `.odttf` font parts,
  unused embedded fonts, the full Word style catalog, exact
  page-number formats, or exact header text/content in some cases.
  In the long-form probe, Google collapsed the style catalog from
  136 styles to 10, and roman/upper-letter page-number formats were
  normalized by render/export behavior.
- **Google Docs -> DOCX** did preserve `docProps/custom.xml` for
  TokenMoulds string properties, used embedded fonts after rewriting
  them as `.ttf`, sections, columns, landscape pages, and
  header/footer references. Practical rule: for Docs, prefer
  `docProps/custom.xml` as the best invisible-ish OOXML metadata
  carrier; do not depend on `customXml/*` or document settings
  sidecars surviving.
- **Google Docs -> DOCX field survival** is narrow. A 2026-05-15
  seed-based field-family probe generated one `w:fldSimple` for each
  Microsoft Word `WdFieldType` family, imported the DOCX into Google
  Docs, and exported it back to DOCX. In body content, only
  `DOCPROPERTY`, `NUMPAGES`, and `TOC` preserved field semantics;
  `TOC` was rewritten into a structured table-of-contents block. In
  headers/footers, `PAGE` and `NUMPAGES` preserved field semantics.
  Google rewrote surviving simple fields to complex
  `w:fldChar`/`w:instrText` runs. Body `PAGE`, body `PAGEREF`, and
  the other body field families flattened to cached text or were
  rewritten without field instructions. Practical rule: treat the
  listed survivors as the Google-compatible DOCX field allow-list and
  keep all other Word field types DOCX-only unless a targeted fixture
  proves otherwise. This probe covered field families, not switch
  permutations.
- **Google Docs -> DOCX real numbering survives as list structure.**
  A 2026-05-15 seed-based probe imported a DOCX with
  `word/numbering.xml`, nested `w:numPr` paragraphs, a restarted list,
  and a legal-style outline. Google export returned all 10 numbered
  paragraphs with `w:numPr` and no literal number prefixes in the
  paragraph text. Google preserved `w:abstractNum`, `w:num`,
  `lowerLetter`, and custom level text such as `Article %1`, but
  normalized the definitions by dropping `w:multiLevelType`,
  `w:isLgl`, and `w:startOverride`, and adding empty placeholder
  levels. Practical rule: represent Google-compatible auto-numbering as
  real list/outline structure, not `AUTONUM`/`SEQ` fields, and treat
  the exported numbering XML as Google-canonical rather than
  byte-stable.
- **Google Docs -> DOCX drawing captions survive with `wps`/`wpg`,
  but not `wpc`.** A 2026-05-15 seed-based probe roundtripped four DOCX
  image/caption carriers. A normal inline `pic:pic` image survived, but
  any caption paragraph stayed separate. A promoted `wps:wsp` shape with
  image `a:blipFill` and `w:txbxContent` caption preserved both the
  image and editable caption text; Google rewrote it as
  `mc:AlternateContent` with a picture fallback. A `wpg:wgp` group
  containing `pic:pic` plus a `wps:wsp` caption textbox also preserved
  the image, group, and caption text, again with a Google-added picture
  fallback and duplicated media parts. A `wpc:wpc` canvas containing a
  picture plus caption textbox was stripped entirely on export. Practical
  rule: promote Google-targeted image captions to `wpg` groups or
  `wps` image-fill shapes, and avoid `wpc` canvas as a carrier.
- **Google Docs -> PDF bitmapifies drawing surfaces at drawing-bound
  resolution.** A 2026-05-15 probe exported `wpg:wgp` image/caption
  drawings to PDF. DOCX export preserved the original 96 x 96 PNG plus
  Google's fallback media, but PDF export embedded a raster image sized
  from the drawing bounds at about 96 px/in. A 1 in drawing became
  96 x 129 px, the same source picture stretched to 4 in became
  384 x 417 px, and one drawing containing both 1x and 4x copies became
  499 x 417 px, matching the combined 5.2 in x 4.35 in extent. Practical
  rule: drawing carriers can preserve editable captions, but print/PDF
  resolution is governed by the physical drawing size and the source
  bitmap's own detail ceiling. A follow-up synthetic resolution-chart
  probe with a 1200 x 1200 Siemens-star/line-pair PNG confirmed the same
  ceiling: DOCX export preserved the 1200 x 1200 source image, while PDF
  export emitted one 499 x 417 raster for the combined same-drawing 1x
  plus 4x group. The 4x crop was sharper at its native size because it
  received 384 x 384 px instead of 96 x 96 px: it had 75.4% near
  black/white pixels versus 39.1% for the 1x crop, and retained visible
  line-pair modulation down to roughly the 8 px source-period patch. This
  is a size-based raster allocation effect, not a high-DPI workaround when
  the final physical image size must remain 1x. A live Google Doc sanity
  check against a normal inline image found the normal picture exported to
  PDF at 2048 x 1226 px, about 452 ppi at its printed size, while two
  `wpg` drawing versions using a 2500 x 1496 source bitmap exported to
  PDF as 433 px wide rasters at 96 ppi. A paired RGB/RGBA PNG probe at
  100 x 100 px and 400 x 400 px, inserted as normal body and header images,
  found no general alpha-channel sharpness win: RGB exported losslessly at
  source dimensions, RGBA exported at the same dimensions as image plus soft
  mask, and DOCX export preserved alpha media. A WPG transform mutation
  confirmed the control variable: multiplying group child coordinates and
  `a:chExt` by 4 while keeping the outer `wp:extent` unchanged left PDF
  drawing rasters at 433 px wide and 96 ppi; scaling the outer `wp:extent`
  and group extent by 1.4 raised them to 606 px wide at 96 ppi and affected
  pagination. A whole-page scale probe doubled page size, margins, normal
  drawing extents, and `wpg` group extents; Google exported an A2 page with
  WPG rasters doubled to about 866 px wide while still marked 96 ppi. Fitting
  that PDF back down to A4 made the WPG rasters report as 192 ppi, but the
  simple Ghostscript post-process recompressed images and stripped PDF tagging.
  The constrained Apps Script PDF extractor can pull these Google-exported
  FlateDecode image XObjects as PNG, including soft-mask RGBA composition; a
  real export probe extracted 400 x 400 alpha and 866 px WPG rasters without
  Poppler. Practical rule: keep print-critical pictures as normal images when
  possible; use drawing carriers when grouped editable caption semantics matter
  more than sharpness, and treat whole-page scaling as a post-export print
  workaround.
- **Google Sheets -> XLSX** does not reliably export
  `docProps/custom.xml`, `customXml/*`, defined-name literal
  metadata, or the core/app docProps in lossy exports. It did
  preserve hidden metadata sheets, visible cells, comments, and a
  defined name pointing at a hidden sheet, though defined-name
  attributes may be normalized. Practical rule: use native workbook
  content carriers, especially a hidden metadata sheet, rather than
  OOXML package sidecars.
- **Google Slides -> PPTX** does not reliably export
  `docProps/custom.xml`, `customXml/*`, shape-name metadata, or the
  `tableStyles` relationship/part. It also normalizes PPTX-authored
  transitions/effects. A 2026-05-08 exhaustive file-level sweep
  (`docs/pptx_oracle/google-slides-animation-roundtrip-2026-05-08.md`)
  covered 98 valid PPTX cases. Only base `p:fade` and `p:push`
  transitions survived as the same transition kind; most other base
  transitions exported as `p:fade`, `p:cover`/`p:pull` exported as
  `p:push`, no `p14:` transition survived as `p14:`, and the
  `p15:prstTrans` smoke case was dropped. For timing, fade entrance
  and exit effects survived as `filter="fade"`, but every other
  entrance/exit `animEffect` filter exported as `filter="fade"` with
  the same direction. Google preserved timing values as millisecond
  integers (`dur="1500"` and the start delay survived), normalized
  `nodeType="clickEffect"` to `nodeType="withEffect"`, remapped shape
  ids, and dropped `filter="image"` opacity emphasis. It did preserve
  speaker notes, hidden slides, off-slide text, shape description/alt
  text, and visible text. Practical rule: use native presentation
  content carriers such as notes, hidden slides, off-slide text, or
  alt text; for imported animations, use PPTX timing XML but restrict
  Google-compatible effects to fade and treat transition speed,
  non-fade effects, p14/p15 transitions, trigger nodeType metadata,
  and original shape ids as non-stable converter details.
- **Google Slides -> PPTX shape effects** are similarly selective.
  A 2026-05-09 DocsServiceApp DrawingML designer smoke imported a
  one-slide PPTX with nine native DrawingML shapes. The source PPTX
  contained `a:gradFill` (9), `a:outerShdw` (6), `a:glow` (6),
  `a:softEdge` (1), and `a:reflection` (1). The Google-exported PPTX
  preserved `a:outerShdw` (6) and `a:reflection` (1), preserved most
  gradients (`a:gradFill` 8), but exported `a:glow` (0) and
  `a:softEdge` (0). Practical rule: for Google-targeted graphics,
  prefer gradients, shadows, and reflection; treat glow and soft-edge
  as non-stable.
- **Google Slides placeholder font settings depend on the carrier, not
  just on the presence of `p:ph`.** A 2026-05-09 generic Google-authored
  template probe exported slide placeholder run fonts (`Courier New`,
  `Georgia`), imported that PPTX back into Google Slides, and exported
  again; the slide placeholder run fonts survived. Layout placeholder
  default fonts also survived, though Google added Arial defaults. A
  generated Figma/SVG PPTX initially lost layout placeholder typefaces
  when defaults were authored only as paragraph-level `a:pPr/a:defRPr`.
  Reauthoring layout placeholder defaults as
  `a:lstStyle/a:lvl1pPr/a:defRPr` made the generated typefaces survive
  Google import/export (`Inter Tight`, `Space Mono`, `Syncopate`, with
  Google adding `Arial`). Google still stripped slide-placeholder
  default style carriers while preserving direct text-run styles.
  Practical rule: for Google-targeted PPTX generation, put master/layout
  placeholder defaults in list-style level defaults and keep visible text
  runs explicitly styled. Do not treat an added `Arial` default as loss
  when the requested typefaces still survive. Full finding:
  `docs/pptx_oracle/google-slides-placeholder-style-roundtrip-2026-05-09.md`.
- **Google Slides native export -> PPTX** uses a different, smaller
  but stable shape. A user-provided native Google Slides export oracle
  (`docs/pptx_oracle/google-slides-native-export-oracle-2026-05-08.md`)
  serialized seven slides as two fade transition forms, two push
  transition forms, and three slides with no transition. Slide timing
  used 12 effect containers: appear/disappear via `p:set`
  `style.visibility`, fade in/out via `p:animEffect filter="fade"`,
  fly/zoom-style effects via `p:anim` over `ppt_x`, `ppt_y`,
  `ppt_w`, and `ppt_h`, and spin via `p:animRot` over `r`. Reimporting
  that Google-exported PPTX and exporting it again preserved the
  transition/effect signatures; the only normalized timing delta was
  an exit visibility-set child delay changing from `0` to `1`.
  A follow-up elaboration probe showed that Google accepts richer
  `p:anim` / `p:animRot` XML but canonicalizes it: arbitrary rotation
  magnitudes/directions and `from`/`to` rotations export as the same
  one-spin `animRot by="-21600000"` shape; far motion, extra keyframes,
  diagonal X+Y, half-size zoom, and overshoot formulas are reduced to
  Google's coarse native fly/zoom shapes.
  Practical rule: for Google-targeted PPTX generation, prefer the
  Google-native export shape instead of broad PowerPoint
  `animEffect` filter vocabulary, and keep motion/zoom/rotation
  parameters canonical.

### Phase 3 — ODF (optional)

Google imports `.odp` / `.odt` / `.ods`. Covered for evidence-ladder
symmetry; deferred from Phase 1/2 to limit blast radius.

## Acceptance Criteria (Phase 1)

1. `src/openxml_audit/gsuite/client.py` exposes the four primitives;
   service account auth works against a real Shared Drive.
2. `tools/oracle/gsuite_roundtrip.py` orchestrates PPTX roundtrip
   and emits `RoundtripObservation` with `loss: set[LossClass]`.
3. `python -m openxml_audit.oracle gsuite FILES...` dispatches correctly.
4. Tests pass without GSuite creds (skip pattern), pass with creds
   when `GSUITE_ORACLE_CREDS` is set.
5. README + CHANGELOG document the one-time GCP setup
   (create project, enable Drive API v3, create service account,
   share Drive folder, drop key in `~/.config/openxml-audit/`).
6. Spec committed.

## Out of Scope (Phase 1)

- DOCX, XLSX, ODF (Phase 2/3).
- **Behavioral / semantic oracle.** This spec is a *file-level*
  oracle: it observes what GSuite emits in OOXML serialization. It
  cannot answer semantic questions about Google's runtime IR — e.g.,
  "is the slide still bound to the layout's placeholder, or is it
  now a disconnected snapshot?" or "do animations actually play in
  the Slides app, or are they merely present in the OOXML?"
  Answering those requires driving the Slides UI in a real browser
  (Playwright) and observing behavior — modifying a master and
  checking whether bound slides update, playing a slideshow and
  observing animation execution, etc. That's a separate spec
  (provisional: 032 — Behavioral GSuite Oracle) with very different
  costs (fragile UI selectors, login flows, slow runs) and is
  explicitly deferred.
- Apps Script Web App as alternative auth model (rejected — service
  account suffices for the oracle's single-principal world).
- Domain-wide delegation / per-user OAuth (the oracle never touches
  user files).
- Loss-class auto-fix or input-side normalization.
- Determinism normalization. Each run may produce slightly different
  exports (timestamps, internal IDs); Phase 2 may add a
  normalize-before-diff pass.
- Drive API quota management beyond honest documentation
  (~20k requests/day free tier ÷ ~5 calls/roundtrip ≈ ~4k files/day
  ceiling).

## Risks

- **GSuite export non-determinism.** Same input may produce
  different bytes across runs. *Mitigation:* `metadata_churn` bucket
  absorbs the dominant source (docProps timestamps); rule-based
  classifier is resilient to spurious noise. Phase 2 may add
  normalization.
- **Mime-type drift.** Google occasionally renames internal mime
  types. *Mitigation:* constants centralized in `gsuite/client.py`;
  rename = one-line fix.
- **Drive API quota exhaustion** during long corpus walks.
  *Mitigation:* documented math; recoverable on next day; Phase 2
  candidate to add batched / resumable runs.
- **Service-account JSON leakage.** Standard Google secret risk.
  *Mitigation:* `.gitignore` enforced; setup docs explicit about
  "never commit this file."
- **Lossy-by-design confused with bug.** A reader could mistake
  "GSuite drops feature X" for a validator finding. *Mitigation:*
  observation report header + module docstring make clear that
  GSuite loss is informational signal about the *target app*, not
  an OOXML conformance failure.
- **Service account can't see user Drive files.** By design — the
  oracle only ever processes files the project itself uploaded. If
  a future use case needs user-side roundtripping, that's a
  different spec with a different auth model.
- **Domain-wide delegation requires Workspace admin.** The setup
  step (Admin Console → Security → API controls → Domain-wide
  Delegation) requires Workspace super-admin rights. *Mitigation:*
  one-time per Workspace; well-documented in the spec's setup
  steps and project README.
- **Impersonated user's quota.** Uploads count against the
  impersonation subject's Drive quota, not the SA's. *Mitigation:*
  oracle deletes both the original upload and the converted Slides
  file in `finally` after each roundtrip; corpus walks should never
  accumulate more than one file's worth at any moment.
