# Architecture Decision Records

This document consolidates architectural decisions for openxml-audit.

This repo owns validation, inspection, and empirical PPTX research tooling.
Converter-side emission and product decisions stay in the sibling
`svg2ooxml` repo.

---

## Mission

### Evidence-Ladder Mission (ADR-001)

`openxml-audit` answers a layered question about Office files: not just
"is this schema-valid?" but "will this survive a roundtrip in the target
app?". The ladder is the organizing principle — validation is the floor
tier; loadability, roundtrip preservation, runtime behavior, and authoring
provenance are the tiers above it. Everything in the repo exists to gather
evidence at one of those tiers, in pure Python, without depending on the
.NET SDK at runtime.

The tiers (`openxml_audit.EvidenceTier`):

1. `schema-valid` — parses against ECMA/OASIS schemas
2. `loadable` — the target app opens the file without repair
3. `roundtrip-preserved` — the app's save does not materially rewrite the intent
4. `slideshow-verified` — runtime behavior in the target app matches intent
5. `ui-authored` — the app itself produced this structure

Decision:

- tier 1 (`schema-valid`) is calibrated against authoritative external
  references (Open XML SDK, ODF specs, fixture contracts), not against any
  one generator repo
- tiers 2-5 are calibrated against curated corpora of target-app-authored
  artifacts (see ADR-002)
- the repo owns validators, package inspection, diffing, lightweight "will
  this work?" capability checks, and the evidence-gathering tools for higher
  tiers
- it does not own file-generation policy for converter products

Consequences:

- consumers can depend on `openxml-audit` without inheriting converter
  assumptions
- evidence drift is handled here, once, across all five tiers, instead of
  being reimplemented in every emitting repo
- the long-term direction is a canonical per-format reference artifact (one
  `.pptx`, one `.docx`, one `.xlsx`) aggregating every feature proven at
  tier `roundtrip-preserved` or above — no such reference document exists
  today, and producing it is the payoff of the corpus work

## PPTX Research

### PPTX Evidence: XML-First and Curated (ADR-002)

ADR-001 establishes the evidence ladder; this ADR records how the **PPTX**
layer of that ladder is organized. PPTX is the first format where empirical
evidence-gathering is needed at scale — animation/timing is the canonical
"schema-valid but PowerPoint silently rewrites it" problem — and the
conventions here are the template for future DOCX and XLSX layers.

Decision:

- durable PPTX evidence lives under `docs/pptx_oracle/`
- reusable PPTX lab tooling lives under `src/openxml_audit/pptx/` with
  entrypoints `openxml-audit-pptx-lab` and `openxml-audit-pptx-timing-oracle`
- PPTX capability findings are registered in
  `src/openxml_audit/pptx/capabilities.py` and reference `EvidenceTier` from
  the format-neutral `openxml_audit.evidence` module
- committed artifacts stay XML-first and curated; full dumps, captures, and
  scratch decks stay out of git unless explicitly promoted
- consumer repos should cite or bridge these artifacts instead of copying the
  corpus into their own docs trees

Consequences:

- PPTX evidence has one durable home tied to the shared ladder
- converter repos stay focused on emission and packaging rather than lab churn
- PPTX research tools can serve multiple consumers, not just `svg2ooxml`
- when DOCX/XLSX evidence work starts, it follows this layout: format-specific
  corpus under `docs/<format>_oracle/`, format-specific capability registry
  under `src/openxml_audit/<format>/capabilities.py`, all tiered via the same
  `EvidenceTier` enum — no new abstractions required

### Oracle Deck Scaffold Layer (ADR-003)

The earliest PPTX oracle builders used `python-pptx` directly to create a
blank deck, place a few shapes, save the package, and then patch authored
`<p:timing>` XML into the resulting slides. That was a pragmatic way to move
fast, but it blurred two different concerns:

- the evidence-bearing XML fragments we are actually trying to prove
- the disposable package scaffold used to hold those fragments

Decision:

- PPTX oracle builders should depend on a first-class scaffold layer under
  `src/openxml_audit/pptx/oracle_deck_scaffold.py`
- builders own the authored fragment semantics; the scaffold layer owns
  package persistence and slide-part patching
- `python-pptx` is treated as the current scaffold implementation, not as the
  architectural source of truth for oracle evidence
- future template-backed or fragment-backed deck scaffolds should slot in
  behind this seam without forcing every oracle builder to change again
- committed PPTX oracle decks may be materialized directly from checked-in
  scaffold package trees under `data/pptx_oracle/scaffolds/`

Consequences:

- duplicated zip-patching logic stays in one place
- the code makes the current compromise explicit instead of letting it leak
  into every builder
- committed oracle evidence can move toward checked-in templates or lower-level
  OOXML package writers without another broad refactor
