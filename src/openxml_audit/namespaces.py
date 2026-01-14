"""Open XML namespace definitions.

Based on ECMA-376 and Microsoft Open XML SDK.
"""

# Content Types namespace
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

# Relationships namespaces
RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
RELATIONSHIPS_METADATA_CORE = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)

# Office Document Relationships
REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_EXTENDED_PROPERTIES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
)
REL_CUSTOM_PROPERTIES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
)
REL_THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"

# PresentationML namespaces
PRESENTATIONML = "http://schemas.openxmlformats.org/presentationml/2006/main"
PRESENTATIONML_STRICT = "http://purl.oclc.org/ooxml/presentationml/main"

# WordprocessingML namespace
WORDPROCESSINGML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# SpreadsheetML namespace
SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# PresentationML relationship types
REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
REL_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
REL_NOTES_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
)
REL_NOTES_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
)
REL_HANDOUT_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
)
REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_PRES_PROPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
)
REL_VIEW_PROPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
)
REL_TABLE_STYLES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
)
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
REL_CUSTOM_XML_PROPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
)
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_STYLES_WITH_EFFECTS = (
    "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
)
REL_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
REL_WEB_SETTINGS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
)
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
REL_FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
REL_SHARED_STRINGS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)
REL_WORKSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
REL_CHARTSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
)
REL_DIALOGSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
)
REL_MACRO_SHEET = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"

# DrawingML namespaces
DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"
DRAWINGML_STRICT = "http://purl.oclc.org/ooxml/drawingml/main"
DRAWINGML_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWINGML_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
DRAWINGML_PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture"
DRAWINGML_SPREADSHEET = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWINGML_WORDPROCESSING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

# Office Document namespaces
OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
OFFICE_DOC_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OFFICE_DOC_MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
OFFICE_DOC_BIBLIOGRAPHY = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
OFFICE_DOC_CUSTOM_XML = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
)

# Core Properties (Dublin Core)
CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
DCMITYPE = "http://purl.org/dc/dcmitype/"
XSI = "http://www.w3.org/2001/XMLSchema-instance"

# Extended Properties (App)
EXTENDED_PROPERTIES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)

# VML (Vector Markup Language)
VML = "urn:schemas-microsoft-com:vml"
VML_OFFICE = "urn:schemas-microsoft-com:office:office"
VML_WORD = "urn:schemas-microsoft-com:office:word"
VML_EXCEL = "urn:schemas-microsoft-com:office:excel"
VML_POWERPOINT = "urn:schemas-microsoft-com:office:powerpoint"

# Markup Compatibility
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Microsoft Office extensions
MS_OFFICE = "http://schemas.microsoft.com/office/2006/metadata/properties"
MS_OFFICE_WORD = "http://schemas.microsoft.com/office/word/2006/wordml"
MS_OFFICE_EXCEL = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
MS_OFFICE_DRAWING = "http://schemas.microsoft.com/office/drawing/2010/main"
MS_OFFICE_POWERPOINT = "http://schemas.microsoft.com/office/powerpoint/2010/main"

# XML standard namespaces
XML = "http://www.w3.org/XML/1998/namespace"
XSD = "http://www.w3.org/2001/XMLSchema"

# Namespace prefix map for lxml
NSMAP = {
    "ct": CONTENT_TYPES,
    "r": RELATIONSHIPS,
    "p": PRESENTATIONML,
    "a": DRAWINGML,
    "pic": DRAWINGML_PICTURE,
    "c": DRAWINGML_CHART,
    "dgm": DRAWINGML_DIAGRAM,
    "mc": MC,
    "v": VML,
    "o": VML_OFFICE,
    "w": WORDPROCESSINGML,
    "x": SPREADSHEETML,
    "wp": DRAWINGML_WORDPROCESSING,
    "dc": DC,
    "dcterms": DCTERMS,
    "xsi": XSI,
}

# Reverse map for looking up prefix by namespace
PREFIX_MAP = {v: k for k, v in NSMAP.items()}


def get_prefix(namespace: str) -> str | None:
    """Get the standard prefix for a namespace URI."""
    return PREFIX_MAP.get(namespace)


def qualify_name(local_name: str, namespace: str) -> str:
    """Create a Clark notation qualified name {namespace}local_name."""
    return f"{{{namespace}}}{local_name}"


def split_qualified_name(qname: str) -> tuple[str | None, str]:
    """Split a Clark notation name into (namespace, local_name)."""
    if qname.startswith("{"):
        ns_end = qname.index("}")
        return qname[1:ns_end], qname[ns_end + 1 :]
    return None, qname
