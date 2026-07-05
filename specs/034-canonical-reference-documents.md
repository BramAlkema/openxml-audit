# Spec: Canonical Reference Documents — Ledger-Generated, Tier-Honest

## Status

Proposed (July 5, 2026). Phase 1 targets the next minor release:
the reference build pipeline (ledger → generator → manifest →
self-validation) shipping as `openxml_audit.reference`, with the
PPTX reference deck generated from committed oracle scaffolds and
DOCX/XLSX references generated as scaffolds-with-index awaiting
their first registered findings.

## Problem

ADR-001 names the payoff of the whole corpus effort: "a canonical
per-format reference artifact (one `.pptx`, one `.docx`, one
`.xlsx`) aggregating every feature proven at tier
`roundtrip-preserved` or above — no such reference document exists
today." Two releases of oracle work later, that is still true.

The ingredients all exist and are already connected to the evidence
ladder:

- `CapabilityFinding` / `EvidenceTier` primitives
  (`src/openxml_audit/evidence/`)
- per-format capability registries
  (`src/openxml_audit/{pptx,docx,xlsx}/capabilities.py` — 6 PPTX
  findings, DOCX/XLSX empty at kickoff)
- committed, self-contained PPTX oracle scaffolds whose slides embed
  authored probe XML (`data/pptx_oracle/scaffolds/{oracle_starter,
  timing_oracle}/`)
- minimal-package builders for DOCX/XLSX
  (`docx/oracle_starter_doc.py`, `xlsx/oracle_starter_book.py`)
- four roundtrip oracles that consume any corpus directory

What is missing is the assembly step: nothing walks the registries,
selects qualifying features, and emits one document per format with
a provenance manifest. Without it, the registries and the oracle
baselines stay parallel artifacts instead of ladder rungs.

## Why This Matters

- **It is the deliverable.** ADR-001 defines the reference documents
  as the payoff; every oracle run and capability registration is
  input to them.
- **It closes the loop.** The generated reference document is itself
  corpus material: run it through the repair oracles and the
  observation either promotes its features up the ladder
  (`loadable` → `roundtrip-preserved`) or falsifies a registry
  claim. Generation → oracle → promotion is the evidence ladder
  operating as a cycle instead of a stack of one-way reports.
- **It makes gaps legible.** A tier-honest builder shows exactly
  which registered features lack emitters and which formats lack
  findings, turning "the registries are thin" from a vibe into a
  status report.

## Non-Goals

- No fabricated tier claims. The builder never asserts evidence; it
  reads what the registries already claim and reproduces it in the
  manifest verbatim.
- No hand-authored reference binaries in git. The reference
  documents are derived artifacts, byte-reproducible from committed
  sources (registries + scaffolds + builders). Only the generator
  and its inputs are committed.
- No new evidence gathering in this spec. Running the oracles on the
  generated references and promoting tiers is follow-up work (it
  needs a macOS + Office session; see Ladder Cycle below).
- No changes to the main `openxml-audit` CLI or to `pyproject.toml`
  entry points in Phase 1 (module is invocable via
  `python -m openxml_audit.reference`).

## Design

### Ledger

`openxml_audit.reference.ledger` aggregates the per-format
capability registries into one cross-format view:

- `TIER_ORDER` fixes the ladder ranking:
  `schema-valid < loadable < roundtrip-preserved <
  slideshow-verified < ui-authored` (ADR-001 order).
- A finding **qualifies** at minimum tier T when any of its
  registered tiers ranks ≥ T. Registered tiers remain sets, not
  cumulative claims; the manifest always carries the exact
  registered tuple so rank-based selection loses no information.
- `LedgerEntry` = (format, finding, emitter availability).

### Emitters

`openxml_audit.reference.emitters` binds capability keys to
document fragments. The binding is per-format:

- **PPTX**: a feature maps to one or more committed scaffold slides
  (`PptxSlideSource(scaffold_name, slide_number)`). The scaffold
  slides already embed the authored `<p:timing>` probes and their
  visual controls — they are the evidence-bearing XML, reused
  as-is. Phase 1 bindings:
  - `pptx.timing.end-condition.time-offset` → `timing_oracle` slide 2
  - `pptx.timing.end-condition.click` → `timing_oracle` slide 3
  - `pptx.timing.repeat-duration` → `timing_oracle` slide 4
  - `pptx.timing.restart` → `timing_oracle` slides 5–6
  - `pptx.anim.effect.entr.fade` / `.wipe` — **no emitter yet**
    (the registered findings describe plain entrance effects; no
    committed scaffold slide exercises exactly that structure).
    They appear in build output as emitter gaps, not silently.
- **DOCX**: a feature maps to a callable returning `<w:body>` block
  elements, appended under a per-feature heading. Registry empty at
  kickoff (matches the empty capability registry).
- **XLSX**: a feature maps to a callable returning `<row>` elements
  for the feature sheet region. Registry empty at kickoff.

### Builders

`openxml_audit.reference.documents` produces, per format:

- the reference document itself:
  - **PPTX**: repackage the scaffold — copy master/layouts/theme/
    props parts verbatim, keep only the selected feature slides
    (renumbered), prepend a generated index slide listing each
    included feature key, summary, and registered tiers; rewrite
    `presentation.xml`, its rels, and `[Content_Types].xml`
    accordingly.
  - **DOCX**: `build_minimal_docx` with a generated body — title,
    per-feature sections from emitters, explicit "no qualifying
    findings" note when the ledger is empty at the requested tier.
  - **XLSX**: `build_minimal_xlsx` with an index sheet — header
    row, one row per feature, note row when empty.
- `<artifact>.manifest.json` — the provenance record: format,
  minimum tier, generator version, and per-feature entries
  (key, summary, registered tiers, inclusion status, location in
  the document, calibration artifacts, constraints, notes), plus
  an `excluded` list with machine-readable reasons
  (`below-minimum-tier`, `no-emitter`).
- **Self-validation**: every built document is validated with
  `OpenXmlValidator` before the build is reported successful. A
  reference document that fails our own tier-1 floor is a build
  error, not an artifact.

Builds are deterministic: no timestamps, stable ordering (ledger
sorted by key), fixed zip member order — same inputs, same bytes.

### CLI

`python -m openxml_audit.reference`:

- `build --format {pptx,docx,xlsx,all} --minimum-tier TIER --out DIR`
  — default tier is `roundtrip-preserved` (the ADR-001 contract).
  Today that produces sparse documents; `--minimum-tier loadable`
  produces the working set. Both are honest; the manifest records
  which contract the artifact was built under.
- `status [--json]` — per-format ledger coverage: findings per
  tier, emitter coverage, qualifying counts at both `loadable` and
  `roundtrip-preserved`. This is the gap report.

### Ladder Cycle (follow-up, not Phase 1)

1. `build --minimum-tier loadable` → reference documents
2. run the repair oracles on the generated documents
3. observations that come back `preserved` justify adding
   `roundtrip-preserved` to the corresponding findings' tier tuples
   (with the baseline JSON recorded in `calibration_artifacts`)
4. rebuild at default tier — the reference document grows

Step 3 is a human-reviewed registry edit on evidence, never
automatic.

## Acceptance

- `python -m openxml_audit.reference build --format all
  --minimum-tier loadable --out DIR` produces three documents +
  three manifests; all three documents pass `OpenXmlValidator`.
- The PPTX reference at `loadable` contains the index slide plus
  the five timing-probe slides; its manifest lists fade/wipe under
  `excluded` with reason `no-emitter`.
- The DOCX/XLSX references build and validate with zero features;
  their manifests say so explicitly.
- `status` reports the fade/wipe emitter gap and the empty
  DOCX/XLSX registries.
- Builds are byte-reproducible run-to-run.
- No changes to `pyproject.toml`, the main CLI, or any existing
  validator behavior.
