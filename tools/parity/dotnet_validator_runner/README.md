# `openxml-validator-runner` — .NET SDK runtime parity tool

Walks an OOXML corpus, invokes the `OpenXmlValidator` from `DocumentFormat.OpenXml` v3.5.1 at every requested `FileFormatVersion`, and emits per-file JSON results. Built for [Spec 013 Open Question 8](../../../specs/013-validator-output-sovereign-gates.md) — the "live SDK runtime as advisory anchor" approach to parity.

## Why this exists

The original parity gate compared our Python validator against expectations *extracted* from SDK test sources by `scripts/corpus/extract_sdk_expectations.py`. Spec 012's `/autoplan` review surfaced multiple landmines in that approach: descriptions templated through `<value>` lose attribute names; `descriptions_by_version` is collected but never written; the manifest became mixed-semantics (some entries SDK-extracted, some manually adjusted to self-parity).

Running the SDK's `OpenXmlValidator` directly against the same corpus eliminates the extraction layer entirely. Same input, two outputs, diff. Plus we get the full validator output (raw attribute names, full XPaths, real error IDs) for every file, not just the test-author's asserted counts.

## Build

```bash
dotnet build tools/parity/dotnet_validator_runner/OpenXmlValidatorRunner.csproj -c Release
```

CI uses a commit-pinned Forgejo `setup-dotnet` action with `dotnet-version: "8.0.x"` (wired in `.forgejo/workflows/calibrate-parity.yml`). The project targets `net8.0`, matching the supported SDK installed by the canonical Forgejo workflow.

## Run

```bash
dotnet tools/parity/dotnet_validator_runner/bin/Release/net6.0/openxml-validator-runner.dll \
  --input-root /tmp/parity-corpus/files \
  --output /tmp/sdk_runtime_snapshot.json \
  --version Office2007 --version Office2010 --version Office2013 --version Office2016 --version Microsoft365
```

If `--version` is omitted, all five default versions run.

## Output schema

```json
{
  "generated_at_utc": "2026-04-29T00:00:00Z",
  "sdk_package_version": "3.5.1.0",
  "validator_versions": ["Office2007", ...],
  "input_root": "/tmp/parity-corpus/files",
  "file_count": 886,
  "files": [
    {
      "source_relpath": "TestFiles/Document.docx",
      "size_bytes": 817192,
      "validations": [
        {
          "version": "Office2007",
          "error_count": 415,
          "errors": [
            {
              "id": "Sch_UndeclaredAttribute",
              "error_type": "Schema",
              "part": "/word/document.xml",
              "path": "/w:document[1]/w:body[1]/...",
              "description": "The 'http://...:firstRow' attribute is not declared."
            }
          ],
          "open_error": null
        }
      ]
    }
  ]
}
```

`open_error` is non-null and `errors` is empty when the OOXML package fails to open (corrupted ZIP, missing parts, etc.). The runner does not abort on per-file failures.

## Comparator

`scripts/parity/diff_sdk_runtime.py` reads this snapshot, runs our Python validator on the same corpus, and diffs the family-key sets per (file, version). Pass `--filter <substring>` to scope to a specific file. Strips XML namespace prefixes (`w:`, `mc:`, `v:`, etc.) from SDK paths before normalizing — necessary because our Python validator emits prefix-less paths and the SDK preserves them.

```bash
python scripts/parity/diff_sdk_runtime.py \
  --sdk-runtime /tmp/sdk_runtime_snapshot.json \
  --files-root /tmp/parity-corpus/files \
  --filter "TestFiles/Document.docx" \
  --only-deltas
```

The comparator is the *intended* consumer; the runner alone is just a snapshot.

## Known artifacts of this proof-of-concept

- **Path normalization needs alignment.** The SDK emits paths with namespace prefixes preserved (`/w:document[1]/...`), our Python validator emits them stripped (`/document[1]/...`). The comparator strips SDK prefixes to match; this works but is a comparator-side hack. A cleaner Spec 013 implementation would either (a) emit prefixes from the Python side too, or (b) normalize both via a shared helper at emission time.
- **Element-indexing discrepancies** between Python and SDK on files with structured document tags (`<w:sdt>`). On `Document.docx` the SDK sees `sdt[5]` and `sdt[11]` where our Python validator sees `sdt[1]`. This is a real divergence and is what the runner is designed to surface.
- **File 886 vs 77 corpus checks.** The runtime snapshot covers all OOXML files in the corpus (886 currently); the parity gate's `manifest.json` scopes to 77 specific assertion checks. Spec 013 needs to decide whether the runtime gate operates over the full corpus or the manifest's subset.
- **No CI wiring yet.** This is a tool, not a gate. Spec 013 owns the policy decision of when/how to wire this into a blocking or advisory CI gate.

## Future work (for Spec 013 proper)

- Decide path-emission alignment between Python and .NET runtime.
- Decide corpus scope (full vs manifest-subset).
- Decide whether to ship a fully runtime-vs-runtime gate (replacing `compare_to_baseline.py`'s frozen-expectations diff) or keep both.
- Possibly delete `scripts/corpus/extract_sdk_expectations.py` and the `expectations[]` lists from `manifest.json` once the runtime gate is sovereign.
