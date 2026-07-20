# Spec: Enforcement Fidelity — One Source of Qualification, One Missing Invariant

## Status

Proposed (July 20, 2026). Extends Spec 009 (`009-bridge-whole-dataset-invariants.md`)
to the schematron bridge, which that spec named in scope but never covered.

Prompted by issue #6 (duplicate `wp:docPr/@id` unreported). The reported
bug was fixed narrowly; the audit behind it found the cause is systemic.

**This spec explicitly recommends against an architectural overhaul.** See
"Why not a rewrite" below.

## Problem

543 of 828 attribute-bearing semantic constraints are shipped, loaded,
converted, and **inert**. They never match an attribute, so they never
reject anything. The test suite reports 948/948 rules converted and is green.

### The measured state

Executed audit over `UNIQUE_ATTRIBUTE` (build the constraint as the running
validator does, synthesise a violating instance with the attribute in the
form the SDK schema data *declares*, run `validate()`):

| | count |
|---|---:|
| live | 135 |
| **dead** | **77** |
| unresolved | 1 |

Bidirectional harness check: `w:` attributes are genuinely namespaced in
instance documents and must classify live — 114 live, 0 dead. VML: 13 live.
The classifier is validated in both directions.

Declarative check (does the constraint look for the attribute in the form
real files write it?) across all attribute-bearing constraints, cross-validated
by returning exactly 77 for `UNIQUE_ATTRIBUTE` and spot-confirmed by
execution on `ATTRIBUTE_VALUE_RANGE`, `ATTRIBUTE_VALUE_LENGTH` and
`ATTRIBUTE_EQUALS` (each: 0 errors with the attribute as real files write
it, 1 error in the form the constraint expects):

| constraint type | total | inert |
|---|---:|---:|
| ATTRIBUTE_VALUE_RANGE | 236 | 224 |
| ATTRIBUTE_VALUE_LENGTH | 191 | 180 |
| UNIQUE_ATTRIBUTE | 213 | 77 |
| ATTRIBUTE_EQUALS | 26 | 23 |
| ATTRIBUTE_NOT_EQUAL | 21 | 19 |
| ATTRIBUTE_VALUE_PATTERN | 22 | 8 |
| ELEMENT_REFERENCE | 23 | 6 |
| ATTRIBUTE_COMPARISON | 6 | 6 |
| ELEMENT_REFERENCE | 22 | 6 |
| **total (resolved)** | **729** | **543** |

**Confidence bound on 543.** Only `UNIQUE_ATTRIBUTE` is *executed*; the rest
is the declarative check. The two agree exactly where they overlap (77), and
three other types were spot-confirmed by execution — but the declarative tier
reports 6 `w:` and 10 `v:` mismatches outside `UNIQUE_ATTRIBUTE`, and both of
those families should be correct today (`w:` is genuinely namespaced, `v:` is
covered by the existing allowlist). Either those are real gaps in the allowlist
or the declarative tier has false positives there. **543 is an estimate with a
known-suspect tail, not a verified count.** Pinning it down requires per-type
violation generators — see Design 2.

### Proximate cause

`codegen/schematron_bridge.py` — `_VML_NAMESPACES`, a two-entry allowlist:

```python
# VML namespaces — attributes are unnamespaced in XML even when the
# schematron data uses the element prefix (e.g. "v:id" → plain "id").
_VML_NAMESPACES = frozenset({...})
```

The comment states the correct general principle and then applies it to VML
only. DrawingML, SpreadsheetML and PresentationML also write unprefixed
attributes, so `wp:id` becomes `{…wordprocessingDrawing}id`, real files carry
a plain `id`, and `validate()` short-circuits on its first line.

This is precisely the failure pattern Spec 009 named: the SDK identifies
something exactly, the bridge substitutes a heuristic, and validation
silently weakens.

### Root cause: three guards, one unasked question

The proximate cause is a one-line allowlist. The reason it survived — and
the reason the next one will too — is that **nothing in the system ever asks
whether a constraint rejects anything.**

| Guard | What it asserts | Why it is blind |
|---|---|---|
| `test_schematron_coverage.py` | `converted == 948`, `skipped_no_constraint == 0` | Measures *conversion*, not effect. A perfectly-formed constraint that can never match counts as covered. |
| Spec 009 invariants (`test_codegen_bridge_invariants.py`) | 21 whole-dataset invariants | Right idea, wrong half. Zero mentions of schematron; imports only `constraint_bridge`. Spec 009 listed `schematron_bridge.py` as normative and never covered it. |
| Parity gate | Output vs SDK v3.4.1 | Advisory/non-blocking, over a ~17-document corpus of mostly-valid files. On a valid file an inert rule and a satisfied rule are indistinguishable. |

Coverage measures structure. Invariants cover the other bridge. Parity
compares silence to silence. The defect is a guard-coverage gap.

## Why not a rewrite

The user question that prompted this spec invited an overhaul. The answer is no.

- **The architecture is sound and already proven.** Data → typed constraints
  → runtime, guarded by whole-dataset invariants, works: it is exactly what
  protects the schema bridge today, where 21 invariants hold the line against
  this same bug class.
- **This project already designed and rejected the overhaul.** Spec 033 /
  ADR-004 proposed replacing the per-rule Python evaluator with compiled
  XSLT, and was shelved when profiling disproved its premise. Reaching for a
  rewrite to fix a missing test would repeat the mistake that spec already
  corrected on evidence.
- **A rewrite does not fix this.** Any reimplementation that reads attribute
  qualification from a hardcoded list instead of the SDK data ships the same
  bug in new code. The defect is in *where the truth comes from*, not in how
  constraints are evaluated.

What is warranted is one deduplication and one new invariant.

## Design

### 1. One source of qualification truth

The SDK schema data already records qualification per attribute, and it is
already in the repo:

```
QName ":id"    → unqualified (no namespace)
QName "w:val"  → qualified with the w prefix
```

Replace `_VML_NAMESPACES` with a resolver that reads this. Two details the
audit surfaced and any implementation must handle:

- **Inheritance.** Attributes are frequently declared on a `BaseClass`, not
  on the element type (`w:CT_FtnEdn/w:endnote` declares none; its base
  `FootnoteEndnoteType` declares `w:id`). Resolving the chain moved the
  audit's unresolved count from 100 to 1.
- **Ambiguity.** An element may declare both forms (`p:sldMasterId` has `id`
  *and* `r:id`). The schematron's own prefix disambiguates: a foreign prefix
  means the qualified attribute; the element's own prefix is the SDK's
  stand-in for unqualified.

**This is a dependency, not an extraction.** An earlier draft of this spec
proposed factoring a shared qualification resolver out of both bridges. That
was wrong: the schema bridge is already correct and has nothing to share.

`SdkAttribute.qname` carries the declaration verbatim, and `.prefix` returns
`None` for the unqualified form; `constraint_bridge._convert_attribute` then
does `ns = namespace_map.get(attr.prefix) if attr.prefix else None`. Four
lines, driven by authoritative data, right all along.

The asymmetry is what the two bridges are *given*:

| | input | qualification |
|---|---|---|
| schema bridge | `SdkAttribute` with `QName` | reads it |
| schematron bridge | `{Context, Test, App}` strings — `@wp:id` from an XPath | had to guess |

The schematron bridge has no attribute metadata at all, which is why it grew
an allowlist. The fix is to give it the metadata: look up the `SdkAttribute`
for (context element, attribute local name) via `schema_loader` and read
`.prefix`. Same authoritative answer the schema bridge already uses, no new
abstraction, and one less place where qualification is decided.

### 2. The missing invariant: executability

Promote the audit harness to a whole-dataset invariant under Spec 009's
existing pattern:

> For every bridged constraint of an attribute-bearing type, synthesise an
> instance that violates it — with the attribute in the form the SDK schema
> data declares — and assert the constraint emits a finding.

Fires → live. Silent on a genuine violation → the bridge lost the rule.

This asks the question no existing guard asks, and it fails loudly on the
whole class rather than on one namespace at a time. The harness exists —
`scripts/audit_constraint_enforcement.py` — and needs promoting into
`test_codegen_bridge_invariants.py`, not writing.

**Phase it by probe correctness, not by ambition.** A violation probe is only
valid if it genuinely violates every constraint of its type:

- **Now: `UNIQUE_ATTRIBUTE`.** Two elements sharing one value violates every
  uniqueness constraint. Validated; passes the harness self-check.
- **Next: per-type generators.** A generic "long non-numeric string" probe was
  tried for the other types and rejected — `ATTRIBUTE_NOT_EQUAL` is only
  violated by a value that *equals* the forbidden one, and
  `ATTRIBUTE_COMPARISON` needs two related attributes. The harness self-check
  caught this (5 `w:` constraints wrongly classified dead), which is the
  clearest evidence that the self-check earns its place.

The self-check is not optional scaffolding. It is the reason this audit can be
trusted, and it must ship with the invariant: any run where a `w:`/`r:`
constraint classifies dead is a broken harness, not a finding.

### 3. Fix the metric

`test_schematron_coverage.py` must stop reporting conversion as coverage.
Coverage becomes *enforced*, with an explicit inert-count assertion that
starts at the post-fix number and is not allowed to grow.

### 4. Staged enablement, validated against oracles

Enabling ~543 dormant constraints will surface findings on files that
validate clean today. Stage by constraint type, smallest first, and gate
each stage on the **roundtrip oracles**, not on the parity gate.

The parity framing is a trap worth naming: these are the SDK's *own*
schematron rules, so enabling them plausibly moves output *toward* SDK
parity — and the parity gate is advisory anyway. The real risk is the
issue #3 lesson: SDK-schema-derived strictness can be wrong about real
applications. Issue #3 shipped four ordering constraints on SDK-derived
reasoning; 396 oracle scenarios later, every non-canonical ordering was a
false positive. A constraint that is live and wrong is worse than one that
is inert.

Per stage: enable → run the oracles and corpus → any finding on a file the
target app opens clean is a false positive and blocks that stage.

## Non-Goals

- **No evaluator rewrite.** Spec 033 settled this; nothing here reopens it.
- **No whole-corpus executability.** The executed invariant starts at
  `UNIQUE_ATTRIBUTE` only — the one type with a probe proven to violate every
  constraint of its type. Other attribute-bearing types follow as their
  generators are written. `CROSS_PART_COUNT`, `OR_CONDITION`, `AND_CONDITION`
  and `CONDITIONAL_VALUE` additionally need multi-part setup and are out of
  scope for this spec.
- **Dead is not the same as a validation gap.** Some inert rules are already
  covered by hand-written validators — duplicate `x:sheet` `name`/`sheetId`
  is caught by `WorkbookValidator` today. Of 76 distinct inert uniqueness
  pairs, 65 have no mention anywhere in `src/`, but that is a grep proxy, not
  proof. Triage before claiming a gap count.
- **No shared-resolver extraction.** Considered and rejected: the schema
  bridge already resolves qualification correctly from `SdkAttribute.qname`,
  so there is no duplicated logic to factor out. The schematron bridge should
  consume `schema_loader`, not share a new component with it.

## Prerequisite: `schema_loader` does not resolve `BaseClass`

Confirmed while validating Design 1. **This blocks pointing the schematron
bridge at `schema_loader`**, which would otherwise inherit the blind spot.

`SdkElementType` stores `base_class` but nothing walks the chain, so an
element whose attributes are all declared on a base type presents an *empty*
attribute set. `w:endnote` resolves to zero attributes while its base
`FootnoteEndnoteType` declares `w:type` and `w:id`.

That empty set then trips the first gate in
`SchemaValidator._should_validate_undeclared_attributes`:

```python
if not constraint.attributes:
    return False  # Skip elements with no declared attributes
```

The gate was written as a safety valve against incomplete SDK metadata. The
metadata is not incomplete — it is unresolved — so the valve fires on a
self-inflicted wound.

**Scale.** Element tags with exactly one type candidate (so the second gate
would pass), zero own attributes, and at least one inherited attribute: **781**.
For comparison, 1520 declare attributes directly and are checked normally.
Affected elements include `a:rPr`, `a:defRPr`, `a:endParaRPr` (19 inherited
attributes each), `a:pPr`/`a:lvl1pPr`…`a:lvl9pPr` (11 each), and `a:clrMap` (12)
— i.e. the elements carrying most DrawingML text formatting.

**Two consequences, one non-consequence:**

- Undeclared-attribute detection is silently off for those 781.
- **1513** inherited attributes carry a `Type` validator that never runs, so
  invalid values in them are never checked.
- Required-attribute checking is *unaffected*: zero inherited attributes are
  marked `Required`.

**Evidence** — end-to-end on one PPTX, each row with a control that fires.
`a:rPr` is inherited-only with one type candidate; `p:cNvPr` and `a:srgbClr`
declare their attributes directly:

| probe | element | result |
|---|---|---|
| baseline | — | 0 findings |
| undeclared attr — control | `a:srgbClr` +`bogusAttr` | **flagged**: "The 'bogusAttr' attribute is not declared." |
| undeclared attr — gap | `a:rPr` +`bogusAttr` | **silent** |
| invalid value — control | `p:cNvPr id="NOT_A_NUMBER"` | **flagged**: "Invalid integer value" |
| invalid value — gap | `a:rPr sz="NOT_A_NUMBER"` | **silent** |
| invalid value — gap | `a:rPr b="NOT_A_BOOL"` | **silent** |

**Fix, with a trap.** Resolve the `BaseClass` chain when building
`SdkElementType.attributes` — but **`ClassName` is not unique across
namespaces**. 385 names collide: `ColorType` is both `a:CT_Color/` (no
attributes) and `x:CT_Color/` (five). A global ClassName index silently
resolves a DrawingML base to a SpreadsheetML type and invents attributes that
do not exist. The first version of this analysis did exactly that and had to be
discarded. Resolution must be scoped to the declaring schema file, falling back
to a global lookup only when the name is globally unambiguous;
`scripts/audit_constraint_enforcement.py:build_attr_index` implements this,
with a cycle guard.

Note this does *not* contradict Design 1: the schema bridge's *qualification*
handling is correct. It is the attribute *set* that is incomplete.

## Acceptance

1. **No regression, and the fix does what it claims.** Re-run the
   executability audit after the change and assert: the 135 currently-live
   `UNIQUE_ATTRIBUTE` constraints stay live, the 77 inert ones flip to live,
   and no constraint moves live → dead. A qualification-logic change could
   silently kill a working `w:`/`r:` constraint; this gate is what proves it
   did not.
2. The bidirectional harness check still passes: `w:` constraints classify
   live, VML constraints classify live.
3. `test_schematron_coverage.py` asserts enforced coverage with an
   inert-count ceiling that cannot grow.
4. The executability invariant runs over the full shipped dataset from
   tracked repo data, no network and no SDK checkout (Spec 009 decision 2).
5. Each enablement stage lands with oracle evidence: zero findings on files
   the target app opens clean.
6. Issue #6's fixtures continue to behave: collision warns, distinct ids
   clean.
7. `BaseClass` resolution lands before the schematron bridge consumes
   `schema_loader`, with the evidence table above as its regression test:
   `a:rPr` must reject `bogusAttr`, `sz="NOT_A_NUMBER"` and `b="NOT_A_BOOL"`
   exactly as `a:srgbClr` and `p:cNvPr` already do.
8. `BaseClass` resolution is namespace-scoped, with a regression test pinning
   a known collision: `ColorType` must resolve to `a:CT_Color/` (0 attributes)
   from a DrawingML type, never to `x:CT_Color/` (5).

## Open Questions

- Do the newly-live constraints change `data/corpus/parity_baseline`
  materially, and does the advisory snapshot need re-extraction?
- Should the unresolved-`BaseClass` defect (below) get its own spec? It sits
  in the schema bridge, not the schematron bridge, but this spec's Design 1
  depends on `schema_loader` being trustworthy.
- How many of the 65 unmentioned inert pairs are genuine validation gaps
  rather than duplicates of hand-written checks?
- What explains the declarative tier's 6 `w:` and 10 `v:` mismatches outside
  `UNIQUE_ATTRIBUTE`? Both families should already be correct. This is the
  first thing per-type generators should settle, because it decides whether
  543 is close to right or materially overstated.
