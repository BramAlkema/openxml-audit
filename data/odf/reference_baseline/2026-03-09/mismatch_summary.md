# ODF Reference Comparison Summary

- Generated at: 2026-03-09T10:52:33.508500+00:00
- Input report: data/odf/reference_baseline/2026-03-09/reference_runs.json
- Tools compared: odf_toolkit, opf

## odf_toolkit
- Samples total: 23
- Samples compared: 23
- Samples skipped: 0
- Reference status counts: {'ok': 23}
- Issue totals: python=26, reference=143, matched=0, only_python=26, only_reference=143
- Mismatch categories: only_python={'package': 15, 'schema': 10, 'semantic': 1}, only_reference={'reference': 25, 'package': 41, 'schema': 63, 'security': 14}

### Top only-python families
- 6x content.xml is missing required office:body element
- 3x Manifest entry '<value>' was not found in package
- 2x Root manifest media-type does not match mimetype ('<value>' != '<value>')
- 2x ODF package is missing required member '<value>'
- 2x ODF package is missing required manifest entry '<value>'
- 1x Missing mimetype entry
- 1x Invalid ODF mimetype value '<value>'
- 1x Missing META-INF/manifest.xml
- 1x Invalid manifest.xml: Opening and ending tag mismatch: file-entry line <n> and manifest, line <n>, column <n> (<string>, line <n>)
- 1x Package member '<value>' is not declared in manifest.xml

### Top only-reference families
- 1x /input/valid_minimal.odt: Info: Generator: null
- 1x /input/valid_minimal.odt/mimetype: Error: The ODF package '<value>' contains a '<value>' file containing '<value>', which differs from the mediatype of the root document '<value>'!
- 1x /input/valid_minimal.odt: Error: The ODF mimetype '<value>' is invalid for the ODF XML Schema document!
- 1x /input/valid_minimal.odt/mimetype: Error: For more information on ODF package conformance see https://docs.oasis-open.org/office/OpenDocument/v1.<n>/os/part2-packages/OpenDocument-v1.<n>-os-part2-packages.html#a_2_2_Package_Conformance
- 1x /input/valid_minimal.odt/mimetype: Error: mimetype is not an ODFMediaTypes mimetype.
- 1x /input/valid_minimal.ods: Info: Generator: null
- 1x /input/valid_minimal.ods/mimetype: Error: The ODF package '<value>' contains a '<value>' file containing '<value>', which differs from the mediatype of the root document '<value>'!
- 1x /input/valid_minimal.ods: Error: The ODF mimetype '<value>' is invalid for the ODF XML Schema document!
- 1x /input/valid_minimal.ods/mimetype: Error: For more information on ODF package conformance see https://docs.oasis-open.org/office/OpenDocument/v1.<n>/os/part2-packages/OpenDocument-v1.<n>-os-part2-packages.html#a_2_2_Package_Conformance
- 1x /input/valid_minimal.ods/mimetype: Error: mimetype is not an ODFMediaTypes mimetype.

## opf
- Samples total: 23
- Samples compared: 23
- Samples skipped: 0
- Reference status counts: {'ok': 23}
- Issue totals: python=26, reference=119, matched=0, only_python=26, only_reference=119
- Mismatch categories: only_python={'package': 15, 'schema': 10, 'semantic': 1}, only_reference={'reference': 89, 'schema': 13, 'package': 13, 'security': 4}

### Top only-python families
- 6x content.xml is missing required office:body element
- 3x Manifest entry '<value>' was not found in package
- 2x Root manifest media-type does not match mimetype ('<value>' != '<value>')
- 2x ODF package is missing required member '<value>'
- 2x ODF package is missing required manifest entry '<value>'
- 1x Missing mimetype entry
- 1x Invalid ODF mimetype value '<value>'
- 1x Missing META-INF/manifest.xml
- 1x Invalid manifest.xml: Opening and ending tag mismatch: file-entry line <n> and manifest, line <n>, column <n> (<string>, line <n>)
- 1x Package member '<value>' is not declared in manifest.xml

### Top only-reference families
- 23x "severity" : "ERROR",
- 23x "severity" : "WARNING",
- 22x NOT VALID, <n> errors, <n> warnings and <n> info messages.
- 1x "title" : "Not a valid XML document. Validation exception at line <n> and column <n>: value of attribute \"manifest:checksum\" is invalid; must be a base64 string.",
- 1x "value" : "value of attribute \"manifest:checksum\" is invalid; must be a base64 string"
- 1x INCOMPLETE encrypted entries are not supported, <n> errors, <n> warnings and <n> info messages.
- 1x APP-<n>: [INFO] Validating /input/invalid_missing_mimetype.odt.
- 1x APP-<n>: [INFO] Validation report for /input/invalid_missing_mimetype.odt.
- 1x "filename" : "/input/invalid_missing_mimetype.odt",
- 1x APP-<n>: [INFO] Validating /input/invalid_invalid_mimetype.odt.

## Cross-tool grouped families

### Top grouped only-python families
- 12x content.xml is missing required office:body element (tools={'odf_toolkit': 6, 'opf': 6})
- 6x Manifest entry '<value>' was not found in package (tools={'odf_toolkit': 3, 'opf': 3})
- 4x Root manifest media-type does not match mimetype ('<value>' != '<value>') (tools={'odf_toolkit': 2, 'opf': 2})
- 4x ODF package is missing required member '<value>' (tools={'odf_toolkit': 2, 'opf': 2})
- 4x ODF package is missing required manifest entry '<value>' (tools={'odf_toolkit': 2, 'opf': 2})
- 2x Missing mimetype entry (tools={'odf_toolkit': 1, 'opf': 1})
- 2x Invalid ODF mimetype value '<value>' (tools={'odf_toolkit': 1, 'opf': 1})
- 2x Missing META-INF/manifest.xml (tools={'odf_toolkit': 1, 'opf': 1})
- 2x Invalid manifest.xml: Opening and ending tag mismatch: file-entry line <n> and manifest, line <n>, column <n> (<string>, line <n>) (tools={'odf_toolkit': 1, 'opf': 1})
- 2x Package member '<value>' is not declared in manifest.xml (tools={'odf_toolkit': 1, 'opf': 1})

### Top grouped only-reference families
- 23x /input/valid_minimal.odt: Info: Generator: null (tools={'odf_toolkit': 23})
- 23x "severity" : "ERROR", (tools={'opf': 23})
- 23x "severity" : "WARNING", (tools={'opf': 23})
- 22x /input/valid_minimal.odt: Error: The ODF mimetype '<value>' is invalid for the ODF XML Schema document! (tools={'odf_toolkit': 22})
- 22x /input/valid_minimal.odt/mimetype: Error: mimetype is not an ODFMediaTypes mimetype. (tools={'odf_toolkit': 22})
- 22x NOT VALID, <n> errors, <n> warnings and <n> info messages. (tools={'opf': 22})
- 20x /input/valid_minimal.odt/mimetype: Error: The ODF package '<value>' contains a '<value>' file containing '<value>', which differs from the mediatype of the root document '<value>'! (tools={'odf_toolkit': 20})
- 19x /input/valid_minimal.odt/mimetype: Error: For more information on ODF package conformance see https://docs.oasis-open.org/office/OpenDocument/v1.<n>/os/part2-packages/OpenDocument-v1.<n>-os-part2-packages.html#a_2_2_Package_Conformance (tools={'odf_toolkit': 19})
- 16x APP-<n>: [INFO] Validating /input/invalid_missing_mimetype.odt. (tools={'opf': 16})
- 16x APP-<n>: [INFO] Validation report for /input/invalid_missing_mimetype.odt. (tools={'opf': 16})
