# WordprocessingML Ordering Research

Spec: [`specs/010-word-compat-element-ordering.md`](../../specs/010-word-compat-element-ordering.md)

The XSD-derived ordering proxy in `src/openxml_audit/word/compat.py` is a
research aid, not a runtime validation rule. The 0.8.0 release retired its
warnings after the Word oracle preserved all 396 tested ordering scenarios
without repair, a dialog, or an XML rewrite. This included issue #3's exact
`tblHeader`/`cantSplit` example and full reversals of the tested sequences.

The retained proxy lets research tools compare two sources:

- **SDK schema metadata** (`Children` list per complex type) — used as a proxy
  for Word's actual tolerance. Cheap, but a proxy: Word's behaviour is
  empirical, not the spec.
- **Empirical mining** of real-world DOCX corpora — used to validate or replace
  the proxy when corpus evidence shows the SDK ordering is too strict.

## Mining tool

[`scripts/mine_word_property_orderings.py`](../../scripts/mine_word_property_orderings.py)
walks a DOCX/DOTX corpus, records every property element's child sequence,
and reports:

1. Distinct orderings observed per property type, with frequency counts
2. Whether each ordering is a valid subsequence of the existing constraint
   table (the SDK proxy)
3. Failure clusters — orderings that disagree with the proxy

```bash
python scripts/mine_word_property_orderings.py /path/to/corpus
python scripts/mine_word_property_orderings.py /path/to/corpus --json out.json
python scripts/mine_word_property_orderings.py /path/to/corpus --exclude archive
```

## Runtime status (0.8.0)

`DocumentValidator` does not execute this proxy. Reintroducing an ordering
warning requires a minimal affected package, exact Word build and platform,
and a versioned oracle result demonstrating a repeatable repair boundary.

Committed oracle baselines live under `tools/oracle/baselines/`:

| Property type | Scenarios preserved | Repaired | Dialogs |
|---|---:|---:|---:|
| `trPr` | 68/68 | 0 | 0 |
| `tblPr` | 93/93 | 0 | 0 |
| `tcPr` | 80/80 | 0 | 0 |
| `sectPr` | 155/155 | 0 | 0 |

## Corpus-mining status (April 28, 2026)

Mined the [TokenMoulds](https://github.com/BramAlkema/TokenMoulds) corpus —
a corporate template generator that emits real, Word-acceptable DOCX:

| Type | Subtrees | Pass SDK proxy | Verdict |
|---|---:|---:|---|
| `tblPr` | 4,847 | 100% | Corpus conforms; Word oracle still preserved all reorders |
| `tcPr` | 28,667 | 100% | Corpus conforms; Word oracle still preserved all reorders |
| `sectPr` | 40 | 100% | Corpus conforms; Word oracle still preserved all reorders |
| `pPr` | 11,420 | 99.78% | Proxy is too strict for observed files |
| `rPr` | 23,754 | 98.59% | Proxy is too strict for observed files |
| `trPr` | 0 | — | No corpus observations |

The 334 `rPr` failures cluster around `color` appearing alongside `sz`/`szCs`
in orders the SDK declares wrong but real Word output emits routinely. This
is direct evidence that the XSD-derived sequence is a proxy, not an oracle,
for that complex type.

## How to extend the research proxy

When the corpus says the SDK proxy holds:

1. Add the canonical ordering to `CONSTRAINT_TABLE` in
   `src/openxml_audit/word/compat.py`. Cite the ECMA-376 section in a
   comment and note the corpus validation result.
2. Add proxy tests in
   `tests/test_word_compat_ordering.py`.
3. Re-run the mining tool on the corpus to confirm zero false positives.

Do not wire the result into runtime validation based on corpus conformance.
Real-world files establish what producers emit, not what Word repairs.

When the corpus says the SDK proxy fails:

1. Run the mining tool with `--json` against the broadest corpus you have.
2. Inspect the empirical-canonical output and the failure clusters.
3. Hand-derive a canonical ordering from the dominant patterns; if any
   pair is genuinely ambiguous in the corpus, omit the more-permissive
   one from the constraint and document the reason in source.
4. Commit the empirical canonical with the source corpus statistics in a
   comment for traceability.

## What the corpus *can't* tell us

The corpus shows what real-world DOCX emitters write, not what Word actually
rejects. To convert any of these findings from "high-confidence proxy" to
"oracle" requires explicit ground truth — files that triggered Word's repair
dialog with their offending property trees preserved. A future report must
include that package plus Word build and platform details.
