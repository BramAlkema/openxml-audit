# Spec 038: Euro-Office Conversion Oracle

## Status

Implemented (August 9, 2026): pinned format model, authenticated client,
packaged dispatcher engine, unit contract tests, and live conversion
baseline. The live server reported 9.3.1.37; upstream 9.3.3 is modeled and
recorded separately rather than mislabeling the snapshot.

## Problem

`openxml-audit` has local application oracles for Microsoft Office and
LibreOffice, plus a Google Workspace import/export oracle. The Opus95 office
path now uses Euro-Office Document Server with the Nextcloud connector, but the
repository had no executable description of that connector's current format
semantics and no way to capture conversion-path evidence.

Calling ODF simply "editable" is misleading. The connector converts ODT, ODS,
and ODP to OOXML before editing; ODG is viewable and convertible but is not an
editable format. The oracle must preserve those distinctions and must not turn
a conversion API check into a claim about the browser editor.

## Pinned Upstream Contract

This implementation is calibrated to:

- Euro-Office/Document Server 9.3.3 (released August 4, 2026)
- ONLYOFFICE Nextcloud connector 11.0.1 (released June 30, 2026)
- `ONLYOFFICE/document-formats` 3.2.0, connector-pinned commit
  `7d7576a3fe2337c30f4c9b40fae70a69dc68ba08`

Primary upstream references:

- <https://github.com/ONLYOFFICE/DocumentServer/releases/tag/v9.3.3>
- <https://github.com/ONLYOFFICE/onlyoffice-nextcloud/releases/tag/v11.0.1>
- <https://github.com/ONLYOFFICE/document-formats/tree/7d7576a3fe2337c30f4c9b40fae70a69dc68ba08>

The format matrix relevant to this validator is:

| Mode | Word | Cell | Slide |
|---|---|---|---|
| Native edit | DOCX, DOCM, DOTX, DOTM | XLSX, XLSM, XLTX, XLTM | PPTX, PPTM, POTX, POTM, PPSX, PPSM |
| Lossy edit | ODT, OTT -> DOCX | ODS, OTS -> XLSX | ODP, OTP -> PPTX |
| View/convert only | - | - | ODG -> PPTX |

`src/openxml_audit/eurooffice/formats.py` is the machine-readable version of
this table. Formats outside it are reported as `unsupported`; the oracle does
not infer capability from a file merely being a ZIP/XML container.

## Client Contract

`EuroOfficeClient` uses only the Python standard library and implements four
operations:

1. `GET /healthcheck`
2. `POST /coauthoring/CommandService.ashx` with `{"c":"version"}`
3. `POST /converter?shardKey=<key>` with the connector-compatible synchronous
   conversion body
4. download of the returned `fileUrl`

When `EUROOFFICE_ORACLE_JWT_SECRET` is set, requests receive both forms used by
connector 11.0.1: the configured header token (Authorization/Bearer by default)
and the signed `token` field in the JSON body. HS256 is implemented with
stdlib `hmac`; no JWT dependency is added.

Security rules:

- secrets have no command-line option;
- error messages omit response bodies and URL query strings;
- reports never contain JWTs, response URLs, or the source URL;
- conversion keys are deterministic SHA-256 prefixes, restricted to the
  Document Server key character/length contract.

Unsigned requests remain possible when the target server has JWT disabled.

## Oracle Semantics

For each supported input, `tools/oracle/eurooffice_conversion_oracle.py`:

1. resolves a URL that Document Server can fetch;
2. calls synchronous conversion using the matrix target;
3. downloads the result;
4. validates the returned OOXML package with `OpenXmlValidator`;
5. for same-format conversion, runs `package_diff.compare_packages` over all
   XML and relationship parts;
6. emits a JSON observation and aggregate outcome counts.

Outcomes:

- `preserved`: same-format output has no canonical XML/relationship diff;
- `rewritten`: same-format output is valid but canonical parts changed;
- `converted`: a cross-format target is valid;
- `unsupported`: the pinned connector matrix has no matching capability;
- `request_failed`, `download_failed`, `invalid_output`,
  `source_unavailable`: operational failures, all producing exit status 1.

Unsupported capability is data, not a run failure. Native OOXML, lossy ODF,
and ODG view-only modes remain separate report fields even when conversion
succeeds.

## Source Reachability

Document Server fetches the input itself. HTTP(S) inputs can therefore be
passed directly. A local input needs `--source-base-url` or
`EUROOFFICE_ORACLE_SOURCE_BASE_URL`; the server must be able to retrieve the
local file's URL-escaped basename from that base URL. The oracle copies the
local bytes only for validation and diffing—it does not start an HTTP server or
upload source files.

## Evidence Boundary

This is a **conversion-path oracle**, not a roundtrip editor oracle. It verifies
Document Server health/version, remote fetch, conversion, artifact download,
and static output validity. It does not exercise:

- Nextcloud discovery or connector configuration;
- browser editor boot and WebSocket/coauthoring paths;
- user interaction, comments, macros, or collaborative edits;
- the Nextcloud callback and persistence of an edited file;
- visual or semantic fidelity after conversion.

A full user-path claim still requires a real browser/editor open-edit-save
check against Nextcloud. The report embeds this limitation in
`evidence_scope` so downstream summaries cannot silently erase it.

## Acceptance

- Exact positive and negative format-matrix tests.
- JWT envelope, health/version, conversion-error, and secret non-leakage tests.
- Network-free oracle tests for native preservation, lossy ODF conversion, ODG
  view-only conversion, unsupported formats, and report scope.
- `eurooffice`, `euro-office`, and `euro` dispatcher routes work from an
  installed wheel.
- A built sdist-to-wheel artifact contains `tools.oracle` and all dispatcher
  engines import outside the checkout.
- Live health, version, and representative OOXML/ODF conversions are recorded
  in `tools/oracle/baselines/eurooffice/2026-08-09.json` without committing
  credentials; failures and version drift remain visible evidence.
