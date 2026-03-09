"""Regression tests for semantic validation behavior."""

from __future__ import annotations

from types import SimpleNamespace

from lxml import etree

from openxml_audit import OpenXmlValidator
from openxml_audit.context import ValidationContext
from openxml_audit.namespaces import MC, OFFICE_DOC_RELATIONSHIPS, WORDPROCESSINGML
from openxml_audit.semantic.relationships import (
    RelationshipExistConstraint,
    validate_part_relationships,
)
from openxml_audit.semantic.validator import SemanticValidator, create_word_semantic_validator


def test_mc_ignorable_empty_is_ignored() -> None:
    element = etree.fromstring(
        f'<a xmlns:mc="{MC}" mc:Ignorable=""/>'.encode()
    )
    context = ValidationContext(max_errors=0)
    validator = SemanticValidator()

    validator._validate_mc_ignorable(element, context)

    assert not any("Ignorable attribute is empty" in error.description for error in context.errors)


def test_word_id_refs_are_deduplicated_by_value() -> None:
    xml = etree.fromstring(

            b'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:commentRangeStart w:id="1"/>'
            b'<w:r><w:commentReference w:id="1"/></w:r>'
            b'<w:commentRangeEnd w:id="1"/>'
            b"</w:p>"

    )
    validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)

    refs = validator._iter_word_id_refs(
        xml, ("commentReference", "commentRangeStart", "commentRangeEnd")
    )

    assert refs == ["1"]


def test_word_color_theme_constraint_not_registered_without_sdk_rules() -> None:
    validator = create_word_semantic_validator(load_sdk_rules=False)

    color_tag = f"{{{WORDPROCESSINGML}}}color"
    assert color_tag not in validator._constraints


def test_relationship_scan_does_not_duplicate_missing_target_errors() -> None:
    rel = SimpleNamespace(
        id="rId1",
        is_external=False,
        resolve_target=lambda _part_uri: "/missing.xml",
    )
    part = SimpleNamespace(relationships=[rel], uri="/word/document.xml")
    context = ValidationContext(
        package=SimpleNamespace(has_part=lambda _target: False),
        max_errors=0,
    )

    valid = validate_part_relationships(part, context)  # type: ignore[arg-type]

    assert valid
    assert not any(error.id == "Sem_RelationshipTargetMissing" for error in context.errors)


def test_relationship_exist_constraint_ignores_empty_relationship_id() -> None:
    element = etree.fromstring(
        (
            f'<a xmlns:r="{OFFICE_DOC_RELATIONSHIPS}" '
            'r:blip=""/>'
        ).encode()
    )
    constraint = RelationshipExistConstraint(attribute="blip", namespace=OFFICE_DOC_RELATIONSHIPS)
    part = SimpleNamespace(
        relationships=SimpleNamespace(get_by_id=lambda _rel_id: None),
    )
    context = ValidationContext(max_errors=0)
    context.set_part(part)

    valid = constraint.validate(element, context)

    assert valid
    assert not any(error.id == "Sem_MissingRelationshipReference" for error in context.errors)
