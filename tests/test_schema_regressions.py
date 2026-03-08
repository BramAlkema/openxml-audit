"""Regression tests for schema validator edge cases."""

from __future__ import annotations

from lxml import etree

from openxml_audit.context import ValidationContext
from openxml_audit.namespaces import DRAWINGML, MC
from openxml_audit.schema.particle import (
    AllParticle,
    ChoiceParticle,
    ElementParticle,
    ParticleType,
    get_validator,
)
from openxml_audit.schema.validator import SchemaValidator, get_constraint_for_tag


def test_choice_particle_rejects_trailing_invalid_children() -> None:
    ns = "urn:test"
    choice = ChoiceParticle(
        children=[
            ElementParticle(ns, "valid"),
        ],
        min_occurs=1,
    )
    children = [
        etree.Element(f"{{{ns}}}valid"),
        etree.Element(f"{{{ns}}}invalid"),
    ]
    context = ValidationContext(max_errors=0)

    validator = get_validator(ParticleType.CHOICE)
    assert validator is not None
    assert not validator.validate(choice, children, context)
    assert any("not a valid choice" in error.description for error in context.errors)


def test_all_particle_rejects_unexpected_children() -> None:
    ns = "urn:test"
    all_particle = AllParticle(
        children=[
            ElementParticle(ns, "expected"),
        ],
    )
    children = [
        etree.Element(f"{{{ns}}}expected"),
        etree.Element(f"{{{ns}}}unexpected"),
    ]
    context = ValidationContext(max_errors=0)

    validator = get_validator(ParticleType.ALL)
    assert validator is not None
    assert not validator.validate(all_particle, children, context)
    assert any("Unexpected element" in error.description for error in context.errors)


def test_alternate_content_prefers_matching_choice_over_fallback() -> None:
    alt = etree.fromstring(
        (
            f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:a="{DRAWINGML}">'
            '<mc:Choice Requires="a"><a:fromChoice/></mc:Choice>'
            '<mc:Fallback><a:fromFallback/></mc:Fallback>'
            "</mc:AlternateContent>"
        ).encode()
    )

    validator = SchemaValidator()
    resolved = validator._resolve_alternate_content(alt)

    assert len(resolved) == 1
    assert resolved[0].tag.endswith("fromChoice")


def test_alternate_content_uses_fallback_when_choice_not_understood() -> None:
    alt = etree.fromstring(
        (
            f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:zz="urn:unknown">'
            '<mc:Choice Requires="zz"><zz:fromChoice/></mc:Choice>'
            '<mc:Fallback><fallback/></mc:Fallback>'
            "</mc:AlternateContent>"
        ).encode()
    )

    validator = SchemaValidator()
    resolved = validator._resolve_alternate_content(alt)

    assert len(resolved) == 1
    assert resolved[0].tag.endswith("fallback")


def test_schema_validator_flags_undeclared_attributes() -> None:
    element = etree.Element(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr",
        attrib={"id": "1", "name": "shape", "unknownAttr": "x"},
    )
    context = ValidationContext(max_errors=0)

    validator = SchemaValidator()
    constraint = get_constraint_for_tag(element.tag, element)
    assert constraint is not None
    validator._validate_attributes(element, constraint, context)

    assert any("attribute is not declared" in error.description for error in context.errors)
