"""Feature-survival probes: does a slide still carry a feature's XML?

Each registered PPTX capability key maps to a namespace-aware XPath
signature that identifies the feature's evidence-bearing structure in
slide XML. Oracles use these after a roundtrip to report which
reference-document features survived (Spec 036): given the reference
manifest's feature → slide map, a probe answers "is the structure this
finding describes still present in the roundtripped slide?".

Signatures are intentionally structural (element/attribute presence),
not byte comparisons — a converting app is allowed to reflow XML; the
question is whether the *feature* survived.
"""

from __future__ import annotations

from lxml import etree

from openxml_audit.namespaces import PRESENTATIONML

__all__ = ["PPTX_FEATURE_SIGNATURES", "probe_slide", "probeable_keys"]

_NSMAP = {"p": PRESENTATIONML}

# Capability key -> XPath returning matching nodes when the feature's
# structure is present. `restart` deliberately excludes the tmRoot
# node, which always carries restart="never".
PPTX_FEATURE_SIGNATURES: dict[str, str] = {
    "pptx.anim.effect.entr.fade": (".//p:animEffect[@transition='in'][@filter='fade']"),
    "pptx.anim.effect.entr.wipe": (
        ".//p:animEffect[@transition='in'][starts-with(@filter, 'wipe')]"
    ),
    "pptx.timing.end-condition.time-offset": (".//p:endCondLst/p:cond[@delay][not(@evt)]"),
    "pptx.timing.end-condition.click": (".//p:endCondLst/p:cond[@evt='onClick']"),
    "pptx.timing.repeat-duration": ".//p:cTn[@repeatDur]",
    "pptx.timing.restart": (".//p:cTn[@restart][not(@nodeType='tmRoot')]"),
}


def probeable_keys() -> tuple[str, ...]:
    """Capability keys that have a survival signature."""
    return tuple(sorted(PPTX_FEATURE_SIGNATURES))


def probe_slide(slide_xml: bytes, key: str) -> bool:
    """True when the feature's signature is present in the slide XML.

    Raises `KeyError` for keys without a signature — callers should
    treat those as "not probeable", not as "did not survive".
    """
    xpath = PPTX_FEATURE_SIGNATURES[key]
    root = etree.fromstring(slide_xml)
    return bool(root.xpath(xpath, namespaces=_NSMAP))
