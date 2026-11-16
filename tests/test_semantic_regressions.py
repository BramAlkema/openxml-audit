"""Regression tests for semantic validation behavior."""

from __future__ import annotations

from types import SimpleNamespace

from lxml import etree

from openxml_audit import OpenXmlValidator
from openxml_audit.context import ValidationContext
from openxml_audit.namespaces import MC, OFFICE_DOC_RELATIONSHIPS, WORDPROCESSINGML
from openxml_audit.semantic.references import UniqueAttributeValueConstraint
from openxml_audit.semantic.relationships import (
    RelationshipExistConstraint,
    validate_part_relationships,
)
from openxml_audit.semantic.validator import SemanticValidator, create_word_semantic_validator

_UNIQUE_NS = "urn:openxml-audit-test"


def _unique_tree(values: list[str]) -> etree._Element:
    items = "".join(f'<t:item t:id="{v}"/>' for v in values)
    return etree.fromstring(f'<t:root xmlns:t="{_UNIQUE_NS}">{items}</t:root>'.encode())


def _unique_errors(context: ValidationContext) -> list[str]:
    return [e.node or "" for e in context.errors if e.id == "Sem_UniqueAttributeValue"]


def test_unique_attribute_reports_each_duplicate_value_once() -> None:
    root = _unique_tree(["dup", "dup", "uniq"])
    constraint = UniqueAttributeValueConstraint(
        attribute="id", namespace=_UNIQUE_NS, element_tag=f"{{{_UNIQUE_NS}}}item"
    )
    context = ValidationContext(max_errors=0)
    context.set_part(SimpleNamespace(uri="/a.xml"))

    items = list(root)
    for item in items:
        constraint.validate(item, context)

    # One error, reported on the first occurrence of the duplicate value.
    assert _unique_errors(context) == ["id"]


def test_unique_attribute_all_distinct_produces_no_error() -> None:
    root = _unique_tree(["a", "b", "c"])
    constraint = UniqueAttributeValueConstraint(
        attribute="id", namespace=_UNIQUE_NS, element_tag=f"{{{_UNIQUE_NS}}}item"
    )
    context = ValidationContext(max_errors=0)
    context.set_part(SimpleNamespace(uri="/a.xml"))

    for item in list(root):
        constraint.validate(item, context)

    assert _unique_errors(context) == []


def test_unique_attribute_index_is_rebuilt_per_part() -> None:
    # The per-part scan is memoized in context.part_scratch keyed only by
    # (attr, tag, case) — no root identity — so set_part() MUST clear it or a
    # second part would reuse the first part's stale index and miss its own
    # duplicates. This test fails if the clear regresses.
    constraint = UniqueAttributeValueConstraint(
        attribute="id", namespace=_UNIQUE_NS, element_tag=f"{{{_UNIQUE_NS}}}item"
    )
    context = ValidationContext(max_errors=0)

    part_a = _unique_tree(["x", "x"])
    context.set_part(SimpleNamespace(uri="/a.xml"))
    for item in list(part_a):
        constraint.validate(item, context)
    assert _unique_errors(context) == ["id"]

    part_b = _unique_tree(["y", "y"])
    context.set_part(SimpleNamespace(uri="/b.xml"))
    for item in list(part_b):
        constraint.validate(item, context)
    # Part B's own duplicate is detected (a second "id" error), not masked by
    # part A's cached index.
    assert _unique_errors(context) == ["id", "id"]


def test_mc_ignorable_empty_is_ignored() -> None:
    element = etree.fromstring(f'<a xmlns:mc="{MC}" mc:Ignorable=""/>'.encode())
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
    element = etree.fromstring((f'<a xmlns:r="{OFFICE_DOC_RELATIONSHIPS}" r:blip=""/>').encode())
    constraint = RelationshipExistConstraint(attribute="blip", namespace=OFFICE_DOC_RELATIONSHIPS)
    part = SimpleNamespace(
        relationships=SimpleNamespace(get_by_id=lambda _rel_id: None),
    )
    context = ValidationContext(max_errors=0)
    context.set_part(part)

    valid = constraint.validate(element, context)

    assert valid
    assert not any(error.id == "Sem_MissingRelationshipReference" for error in context.errors)
