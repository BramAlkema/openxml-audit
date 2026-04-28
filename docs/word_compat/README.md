# Word Compatibility Element Ordering — Empirical Mining

Spec: [`specs/010-word-compat-element-ordering.md`](../../specs/010-word-compat-element-ordering.md)

The Word compatibility ordering check (in `src/openxml_audit/word/compat.py`)
flags child element reorderings inside WordprocessingML property complex types
that trigger Word's "unreadable content" repair dialog despite the .NET Open
XML SDK accepting the same files. The canonical orderings come from two
sources:

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

## Validation status (April 28, 2026)

Mined the [TokenMoulds](https://github.com/BramAlkema/TokenMoulds) corpus —
a corporate template generator that emits real, Word-acceptable DOCX:

| Type | Subtrees | Pass SDK proxy | Verdict |
|---|---:|---:|---|
| `tblPr` | 4,847 | 100% | Proxy holds — shipped Phase 2 |
| `tcPr` | 28,667 | 100% | Proxy holds — shipped Phase 2 |
| `sectPr` | 40 | 100% | Proxy holds (small sample) — shipped Phase 2 |
| `pPr` | 11,420 | 99.78% | Borderline; shipping deferred to Phase 3 |
| `rPr` | 23,754 | 98.59% | **Proxy too strict; Phase 3 needs empirical canonical** |
| `trPr` | 0 | — | No corpus data; Phase 1 rests on issue #3 anecdote |

The 334 `rPr` failures cluster around `color` appearing alongside `sz`/`szCs`
in orders the SDK declares wrong but real Word output emits routinely. This
is direct evidence that the XSD-derived sequence is a proxy, not an oracle,
for that complex type.

## How to add a constraint type

When the corpus says the SDK proxy holds:

1. Add the canonical ordering to `CONSTRAINT_TABLE` in
   `src/openxml_audit/word/compat.py`. Cite the ECMA-376 section in a
   comment and note the corpus validation result.
2. Add tests mirroring the Phase 1/2 pattern in
   `tests/test_word_compat_ordering.py`.
3. Re-run the mining tool on the corpus to confirm zero false positives.

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
dialog with their offending property trees preserved. Issue #3 is the place
to collect that data.
