"""Regression tests for schema validator edge cases."""

from __future__ import annotations

from lxml import etree

import openxml_audit.schema.validator as schema_validator_module
from openxml_audit.context import ValidationContext
from openxml_audit.errors import FileFormat
from openxml_audit.namespaces import DRAWINGML, DRAWINGML_CHART, MC, WORDPROCESSINGML
from openxml_audit.schema.particle import (
    AllParticle,
    ChoiceParticle,
    ElementParticle,
    ParticleType,
    SequenceParticle,
    get_validator,
)
from openxml_audit.schema.validator import SchemaValidator, get_constraint_for_tag
from tests.fixture_loader import load_fixture_text


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


def test_choice_particle_with_optional_branches_allows_empty_children() -> None:
    ns = "urn:test"
    choice = ChoiceParticle(
        children=[
            ElementParticle(ns, "a", min_occurs=0, max_occurs=1),
            ElementParticle(ns, "b", min_occurs=0, max_occurs=1),
        ],
        min_occurs=1,
        max_occurs=1,
    )
    context = ValidationContext(max_errors=0)

    validator = get_validator(ParticleType.CHOICE)
    assert validator is not None
    assert validator.validate(choice, [], context)
    assert not context.errors


def test_choice_particle_counts_composite_branch_as_single_occurrence() -> None:
    ns = "urn:test"
    choice = ChoiceParticle(
        children=[
            SequenceParticle(
                children=[
                    ElementParticle(ns, "a", min_occurs=1, max_occurs=1),
                    ElementParticle(ns, "b", min_occurs=1, max_occurs=1),
                ],
                min_occurs=1,
                max_occurs=1,
            ),
            ElementParticle(ns, "b", min_occurs=1, max_occurs=1),
        ],
        min_occurs=1,
        max_occurs=1,
    )
    children = [
        etree.Element(f"{{{ns}}}a"),
        etree.Element(f"{{{ns}}}b"),
    ]
    context = ValidationContext(max_errors=0)

    validator = get_validator(ParticleType.CHOICE)
    assert validator is not None
    assert validator.validate(choice, children, context)
    assert not any("Choice allows at most" in error.description for error in context.errors)


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
            "<mc:Fallback><a:fromFallback/></mc:Fallback>"
            "</mc:AlternateContent>"
        ).encode()
    )

    validator = SchemaValidator()
    resolved = validator._resolve_alternate_content(alt, FileFormat.OFFICE_2019)

    assert len(resolved) == 1
    assert resolved[0].tag.endswith("fromChoice")


def test_alternate_content_uses_fallback_when_choice_not_understood() -> None:
    alt = etree.fromstring(
        (
            f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:zz="urn:unknown">'
            '<mc:Choice Requires="zz"><zz:fromChoice/></mc:Choice>'
            "<mc:Fallback><fallback/></mc:Fallback>"
            "</mc:AlternateContent>"
        ).encode()
    )

    validator = SchemaValidator()
    resolved = validator._resolve_alternate_content(alt, FileFormat.OFFICE_2019)

    assert len(resolved) == 1
    assert resolved[0].tag.endswith("fallback")


def test_schema_validator_ignores_mc_ignorable_extension_children() -> None:
    run = etree.fromstring(
        (
            f'<w:r xmlns:w="{WORDPROCESSINGML}" xmlns:mc="{MC}" '
            'xmlns:w16symex="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
            'mc:Ignorable="w16symex">'
            '<w16symex:sym w16symex:font="Wingdings" w16symex:char="F03A"/>'
            "</w:r>"
        ).encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(run, context)

    assert not any("Unexpected element 'sym'" in error.description for error in context.errors)


def test_schema_validator_ignores_a_ext_extension_entries_in_fallback_mode(
    monkeypatch,
) -> None:
    ext_lst = etree.fromstring(
        (f'<a:extLst xmlns:a="{DRAWINGML}"><a:ext uri="{{123}}"/></a:extLst>').encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    monkeypatch.setattr(schema_validator_module, "_HAS_SDK_CONSTRAINTS", False)
    validator._validate_element(ext_lst, context)

    assert not any("Required attribute 'cx'" in error.description for error in context.errors)
    assert not any("Required attribute 'cy'" in error.description for error in context.errors)
    assert not context.errors


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


def test_schema_validator_selects_supplemental_font_constraint_for_a_font() -> None:
    font = etree.fromstring(load_fixture_text("schema", "supplemental_font.xml").encode("utf-8"))
    context = ValidationContext(max_errors=0)

    validator = SchemaValidator()
    constraint = get_constraint_for_tag(font.tag, font)
    assert constraint is not None

    declared = {attr.local_name for attr in constraint.attributes}
    assert {"script", "typeface"}.issubset(declared)

    validator._validate_element(font, context)
    assert not any("Required element 'latin'" in error.description for error in context.errors)
    assert not any("Required element 'ea'" in error.description for error in context.errors)
    assert not any("Required element 'cs'" in error.description for error in context.errors)


def test_word_cols_allows_empty_col_children() -> None:
    cols = etree.fromstring(
        (
            f'<w:cols xmlns:w="{WORDPROCESSINGML}" w:space="708" w:num="1" w:equalWidth="1"/>'
        ).encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(cols, context)

    assert not any("Required element 'col'" in error.description for error in context.errors)


def test_word_sdt_end_pr_allows_empty_content() -> None:
    sdt_end_pr = etree.fromstring(f'<w:sdtEndPr xmlns:w="{WORDPROCESSINGML}"/>'.encode())
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(sdt_end_pr, context)

    assert not any(
        "Required choice element is missing" in error.description for error in context.errors
    )


def test_word_fld_char_allows_empty_optional_choice_content() -> None:
    fld_char = etree.fromstring(
        (f'<w:fldChar xmlns:w="{WORDPROCESSINGML}" w:fldCharType="begin"/>').encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(fld_char, context)

    assert not any(
        "Required choice element is missing" in error.description for error in context.errors
    )


def test_spreadsheet_cell_style_index_is_not_treated_as_boolean() -> None:
    cell = etree.fromstring(
        b'<x:c xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A1" s="4"/>'
    )
    row = etree.Element("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row")
    row.append(cell)
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(row, context)

    assert not any("Invalid boolean value" in error.description for error in context.errors)


def test_word_footnotes_allows_multiple_footnote_children() -> None:
    footnotes = etree.fromstring(
        (
            f'<w:footnotes xmlns:w="{WORDPROCESSINGML}">'
            '<w:footnote w:id="1"/><w:footnote w:id="2"/>'
            "</w:footnotes>"
        ).encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(footnotes, context)

    assert not any("Unexpected element 'footnote'" in error.description for error in context.errors)


def test_word_settings_allows_do_not_embed_smart_tags_element() -> None:
    settings = etree.fromstring(
        (f'<w:settings xmlns:w="{WORDPROCESSINGML}"><w:doNotEmbedSmartTags/></w:settings>').encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(settings, context)

    assert not any(
        "Unexpected element 'doNotEmbedSmartTags'" in error.description for error in context.errors
    )


def test_word_settings_rejects_do_not_embed_smart_tags_for_office2010() -> None:
    settings = etree.fromstring(
        (f'<w:settings xmlns:w="{WORDPROCESSINGML}"><w:doNotEmbedSmartTags/></w:settings>').encode()
    )
    context = ValidationContext(file_format=FileFormat.OFFICE_2010, max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(settings, context)

    assert any(
        "Unexpected element 'doNotEmbedSmartTags'" in error.description for error in context.errors
    )


def test_chart_dlbls_allows_multiple_optional_boolean_children() -> None:
    d_lbls = etree.fromstring(
        (
            f'<c:dLbls xmlns:c="{DRAWINGML_CHART}">'
            '<c:showLegendKey val="0"/>'
            '<c:showVal val="1"/>'
            '<c:showCatName val="0"/>'
            '<c:showSerName val="0"/>'
            '<c:showPercent val="0"/>'
            '<c:showBubbleSize val="0"/>'
            "</c:dLbls>"
        ).encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(d_lbls, context)

    assert not any("Unexpected element 'showVal'" in error.description for error in context.errors)
    assert not any(
        "Unexpected element 'showLegendKey'" in error.description for error in context.errors
    )


def test_chart_style_and_overlap_value_ranges_match_sdk_types() -> None:
    chart_space = etree.fromstring(
        (
            f'<c:chartSpace xmlns:c="{DRAWINGML_CHART}">'
            '<c:style val="134"/>'
            "<c:chart><c:plotArea><c:barChart>"
            '<c:overlap val="-27"/>'
            "</c:barChart></c:plotArea></c:chart>"
            "</c:chartSpace>"
        ).encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SchemaValidator()

    validator._validate_element(chart_space, context)

    assert not any("Value 134 exceeds maximum 127" in error.description for error in context.errors)
    assert not any(
        "Value -27 is less than minimum 0" in error.description for error in context.errors
    )


def test_spreadsheet_control_rejects_office2010_only_child_in_office2007() -> None:
    control = etree.fromstring(
        (
            '<x:control '
            'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'shapeId="1" r:id="rId1">'
            "<x:controlPr/>"
            "</x:control>"
        ).encode()
    )
    context = ValidationContext(max_errors=0, file_format=FileFormat.OFFICE_2007)
    validator = SchemaValidator()

    validator._validate_element(control, context)

    assert any("Unexpected element 'controlPr'" in error.description for error in context.errors)
