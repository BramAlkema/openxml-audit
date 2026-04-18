# PowerPoint Oracle Fixtures

This directory keeps a small curated subset of extracted PowerPoint timing
oracles in git. Full extraction dumps should stay outside the repository, for
example in local scratch storage, CI artifacts, or release assets.

The committed subset is intentionally small:

- `selected/cloudpresentationpack-azurefunc-animation/slide1`
- `selected/examples-pptx-slide12/slide12`
- `selected/powerpoint-school-slide4/slide4`

Each fixture keeps:

- `timing.raw.xml` - the saved PowerPoint timing XML, used as the source of
  truth
- `timing.normalized.xml` - normalized IDs for stable review and mining
- `summary.json` - extracted effect-pattern metadata

Methodology and ownership:

- `ssot.md`
- `powerpoint-animation-msft-source-map.md`

Empirical findings and fixture notes:

- `animation-oracle-empirical-findings.md`
- `cloudpresentationpack.md`
- `examples-pptx.md`
- `powerpoint-sample-selected.md`

Open research backlog:

- `powerpoint-obscure-animation-backlog.md`

When a larger oracle is needed, regenerate it with
`openxml-audit-pptx-lab oracle` into a temporary or external artifact path
instead of committing the full dump.
