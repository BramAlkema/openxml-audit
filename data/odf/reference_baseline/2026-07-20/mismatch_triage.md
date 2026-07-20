# ODF Mismatch Triage

- Generated at: 2026-07-20T14:41:34.561989+00:00
- Compare input: reports/odf/reference_runs.json
- Tools: odf_toolkit, opf
- Sample count: 126
- Sample categories: {'chart': 4, 'drawing': 7, 'font': 2, 'forms': 4, 'foundation': 3, 'metadata': 3, 'package': 12, 'presentation': 16, 'schema': 6, 'security': 11, 'semantic': 1, 'spreadsheet': 14, 'style': 18, 'text': 15, 'version': 10}
- Sample profiles: {'invalid': 109, 'valid': 17}
- Python issue categories: {'package': 17, 'schema': 12, 'semantic': 112, 'security': 1}

## Actionable Summary
- Prioritize top grouped only-python families for false-positive triage.
- Prioritize top grouped only-reference families for missing-rule coverage.

## odf_toolkit
- Samples compared/skipped: 126 / 0
- Reference status counts: {'ok': 126}
- Mismatch categories: only_python={'package': 17, 'schema': 12, 'semantic': 112, 'security': 1}, only_reference={'reference': 83, 'package': 64, 'schema': 128, 'security': 64}

### Top only-python families
- 8x content.xml contains text:style-name references but styles.xml is not declared in manifest.xml
- 7x Automatic style '<value>' is declared but never referenced
- 6x content.xml is missing required office:body element
- 5x Manifest entry '<value>' was not found in package
- 4x Shape '<value>' (frame) is missing position/size attributes: x, y, width, height
- 3x Frame child references '<value>' which is not present in the package
- 2x Image href '<value>' does not resolve in package
- 2x Shape '<value>' (rect) is missing position/size attributes: x, y, width, height
- 2x Root manifest media-type does not match mimetype ('<value>' != '<value>')
- 2x Manifest media-type for '<value>' should be '<value>' (found '<value>')

### Top only-reference families
- 2x /input/invalid_form_control_id_duplicate.odt/content.xml[<n>,<n>]: Error: element "form:text" is missing "id" attribute
- 1x /input/invalid_aux_declared_missing_styles.odt: Info: Generator: null
- 1x /input/invalid_aux_declared_missing_styles.odt/mimetype: Error: The ODF package '<value>' contains a '<value>' file containing '<value>', which differs from the mediatype of the root document '<value>'!
- 1x /input/invalid_aux_declared_missing_styles.odt/META-INF/manifest.xml: Error: The file '<value>' shall not be listed in the '<value>' file as it does not exist in the ODF package '<value>'!
- 1x /input/invalid_aux_declared_missing_styles.odt: Error: The ODF mimetype '<value>' is invalid for the ODF XML Schema document!
- 1x /input/invalid_aux_declared_missing_styles.odt/mimetype: Error: For more information on ODF package conformance see https://docs.oasis-open.org/office/OpenDocument/v1.<n>/os/part2-packages/OpenDocument-v1.<n>-os-part2-packages.html#a_2_2_Package_Conformance
- 1x /input/invalid_aux_declared_missing_styles.odt/mimetype: Error: mimetype is not an ODFMediaTypes mimetype.
- 1x /input/invalid_aux_invalid_styles_xml.odt: Info: Generator: null
- 1x /input/invalid_aux_invalid_styles_xml.odt/mimetype: Error: The ODF package '<value>' contains a '<value>' file containing '<value>', which differs from the mediatype of the root document '<value>'!
- 1x /input/invalid_aux_invalid_styles_xml.odt: Error: The ODF mimetype '<value>' is invalid for the ODF XML Schema document!

## opf
- Samples compared/skipped: 126 / 0
- Reference status counts: {'ok': 126}
- Mismatch categories: only_python={'package': 17, 'schema': 12, 'semantic': 112, 'security': 1}, only_reference={'reference': 613, 'schema': 28, 'package': 19, 'security': 29}

### Top only-python families
- 8x content.xml contains text:style-name references but styles.xml is not declared in manifest.xml
- 7x Automatic style '<value>' is declared but never referenced
- 6x content.xml is missing required office:body element
- 5x Manifest entry '<value>' was not found in package
- 4x Shape '<value>' (frame) is missing position/size attributes: x, y, width, height
- 3x Frame child references '<value>' which is not present in the package
- 2x Image href '<value>' does not resolve in package
- 2x Shape '<value>' (rect) is missing position/size attributes: x, y, width, height
- 2x Root manifest media-type does not match mimetype ('<value>' != '<value>')
- 2x Manifest media-type for '<value>' should be '<value>' (found '<value>')

### Top only-reference families
- 126x "severity" : "WARNING",
- 80x "severity" : "ERROR",
- 75x NOT VALID, <n> errors, <n> warnings and <n> info messages.
- 46x VALID, no errors, <n> warnings found and <n> info messages.
- 5x INCOMPLETE encrypted entries are not supported, <n> errors, <n> warnings and <n> info messages.
- 4x "title" : "Not a valid XML document. Validation exception at line <n> and column <n>: value of attribute \"manifest:checksum\" is invalid; must be a base64 string.",
- 4x "value" : "value of attribute \"manifest:checksum\" is invalid; must be a base64 string"
- 2x "title" : "Not a valid XML document. Validation exception at line <n> and column <n>: value of attribute \"table:target-range-address\" is invalid; must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+$?[<n>-<n>]+(:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+$?[<n>-<n>]+)?\", must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[<n>-<n>]+:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[<n>-<n>]+\" or must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+\".",
- 2x "value" : "value of attribute \"table:target-range-address\" is invalid; must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+$?[<n>-<n>]+(:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+$?[<n>-<n>]+)?\", must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[<n>-<n>]+:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[<n>-<n>]+\" or must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+\""
- 2x "title" : "Not a valid XML document. Validation exception at line <n> and column <n>: value of attribute \"table:cell-range-address\" is invalid; must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+$?[<n>-<n>]+(:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+$?[<n>-<n>]+)?\", must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[<n>-<n>]+:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[<n>-<n>]+\" or must be a string matching the regular expression \"($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+:($?([^\\. '<value>'([^'<value>''<value>'))?\\.$?[A-Z]+\".",

## Top grouped only-python families
- 16x content.xml contains text:style-name references but styles.xml is not declared in manifest.xml (tools={'odf_toolkit': 8, 'opf': 8})
- 14x Automatic style '<value>' is declared but never referenced (tools={'odf_toolkit': 7, 'opf': 7})
- 12x content.xml is missing required office:body element (tools={'odf_toolkit': 6, 'opf': 6})
- 10x Manifest entry '<value>' was not found in package (tools={'odf_toolkit': 5, 'opf': 5})
- 8x Shape '<value>' (frame) is missing position/size attributes: x, y, width, height (tools={'odf_toolkit': 4, 'opf': 4})
- 6x Frame child references '<value>' which is not present in the package (tools={'odf_toolkit': 3, 'opf': 3})
- 4x Image href '<value>' does not resolve in package (tools={'odf_toolkit': 2, 'opf': 2})
- 4x Shape '<value>' (rect) is missing position/size attributes: x, y, width, height (tools={'odf_toolkit': 2, 'opf': 2})
- 4x Root manifest media-type does not match mimetype ('<value>' != '<value>') (tools={'odf_toolkit': 2, 'opf': 2})
- 4x Manifest media-type for '<value>' should be '<value>' (found '<value>') (tools={'odf_toolkit': 2, 'opf': 2})

## Top grouped only-reference families
- 126x "severity" : "WARNING", (tools={'opf': 126})
- 109x APP-<n>: [INFO] Validating /input/invalid_aux_declared_missing_styles.odt. (tools={'opf': 109})
- 109x APP-<n>: [INFO] Validation report for /input/invalid_aux_declared_missing_styles.odt. (tools={'opf': 109})
- 108x "filename" : "/input/invalid_aux_declared_missing_styles.odt", (tools={'opf': 108})
- 81x /input/invalid_aux_declared_missing_styles.odt: Info: Generator: null (tools={'odf_toolkit': 81})
- 80x "severity" : "ERROR", (tools={'opf': 80})
- 75x NOT VALID, <n> errors, <n> warnings and <n> info messages. (tools={'opf': 75})
- 46x VALID, no errors, <n> warnings found and <n> info messages. (tools={'opf': 46})
- 38x /input/invalid_aux_declared_missing_styles.odt: Error: The ODF mimetype '<value>' is invalid for the ODF XML Schema document! (tools={'odf_toolkit': 38})
- 38x /input/invalid_aux_declared_missing_styles.odt/mimetype: Error: mimetype is not an ODFMediaTypes mimetype. (tools={'odf_toolkit': 38})
