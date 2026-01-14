"""Bridge between SDK schematron rules and semantic constraints.

Converts ParsedSchematron rules into SemanticConstraint instances.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from openxml_audit.codegen.schema_loader import get_registry as get_schema_registry
from openxml_audit.codegen.schematron_loader import (
    ParsedSchematron,
    SchematronType,
    get_registry as get_schematron_registry,
)
from openxml_audit.semantic.attributes import (
    AttributeMinMaxConstraint,
    AttributeValueLengthConstraint,
    AttributeValuePatternConstraint,
    SemanticConstraint,
)
from openxml_audit.semantic.constraints import (
    AndConstraint,
    AttributeEqualsConstraint,
    AttributeNotEqualConstraint,
    AttributesPresentConstraint,
    ConditionalConstraint,
    CrossPartCountConstraint,
    OrConstraint,
)
from openxml_audit.semantic.constraints.equality import AttributeComparisonConstraint
from openxml_audit.semantic.references import UniqueAttributeValueConstraint
from openxml_audit.semantic.relationships import RelationshipTypeConstraint

if TYPE_CHECKING:
    pass


def _get_namespace_map() -> dict[str, str]:
    """Get prefix -> namespace URI map."""
    registry = get_schema_registry()
    registry.load()
    return dict(registry._prefixes)


def _resolve_context_to_tag(context: str, namespace_map: dict[str, str]) -> str | None:
    """Convert schematron context to Clark notation tag.

    Args:
        context: Schematron context like "p:sld" or "x:sheet"
        namespace_map: Prefix to namespace URI mapping

    Returns:
        Clark notation like "{namespace}localname" or None if unresolvable
    """
    if ":" in context:
        prefix, local_name = context.split(":", 1)
        ns = namespace_map.get(prefix)
        if ns:
            return f"{{{ns}}}{local_name}"
    return None


def _split_attribute_name(
    attr: str, namespace_map: dict[str, str]
) -> tuple[str, str | None]:
    """Extract local name + namespace from attribute reference.

    Args:
        attr: Attribute like "x:windowWidth" or "id"
        namespace_map: Prefix to namespace URI map

    Returns:
        (local_name, namespace_uri)
    """
    if ":" in attr:
        prefix, local_name = attr.split(":", 1)
        return local_name, namespace_map.get(prefix)
    return attr, None


def _get_attribute_local_name(attr: str) -> str:
    """Extract local name from attribute reference."""
    if ":" in attr:
        return attr.split(":", 1)[1]
    return attr


def _convert_xpath_pattern(pattern: str) -> str:
    """Convert XPath regex pattern to Python regex (best effort).

    XPath regex uses different syntax for some constructs.
    This handles common cases.

    Args:
        pattern: XPath regex pattern

    Returns:
        Python-compatible regex pattern

    Raises:
        ValueError: If pattern cannot be converted
    """
    import re

    result = pattern

    # XPath uses \p{L} for Unicode letters, Python uses different syntax
    # For now, replace with approximate equivalents
    result = re.sub(r'\\p\{L\}', r'\\w', result)
    result = re.sub(r'\\p\{N\}', r'\\d', result)
    result = re.sub(r'\\p\{[^}]+\}', r'.', result)  # Fallback for other Unicode categories

    # XPath uses \i for XML initial name char, \c for XML name char
    result = re.sub(r'\\i', r'[a-zA-Z_:]', result)
    result = re.sub(r'\\c', r'[a-zA-Z0-9_:.-]', result)

    # Verify the pattern compiles
    re.compile(result)

    return result


def create_constraint_from_schematron(
    rule: ParsedSchematron,
    namespace_map: dict[str, str] | None = None,
) -> SemanticConstraint | None:
    """Convert a parsed schematron rule to a semantic constraint.

    Args:
        rule: The parsed schematron rule.
        namespace_map: Optional prefix->namespace map. If None, will be loaded.

    Returns:
        SemanticConstraint instance, or None if rule cannot be converted.
    """
    if namespace_map is None:
        namespace_map = _get_namespace_map()

    attr_local = _get_attribute_local_name(rule.attribute) if rule.attribute else None
    attr_namespace = None
    if rule.attribute:
        attr_local, attr_namespace = _split_attribute_name(rule.attribute, namespace_map)

    match rule.rule_type:
        case SchematronType.ATTRIBUTE_VALUE_RANGE:
            if attr_local is None:
                return None
            return AttributeMinMaxConstraint(
                attribute=attr_local,
                min_value=rule.min_value,
                max_value=rule.max_value,
                namespace=attr_namespace,
            )

        case SchematronType.ATTRIBUTE_VALUE_LENGTH:
            if attr_local is None:
                return None
            return AttributeValueLengthConstraint(
                attribute=attr_local,
                min_length=rule.min_length,
                max_length=rule.max_length,
                namespace=attr_namespace,
            )

        case SchematronType.ATTRIBUTE_VALUE_PATTERN:
            if attr_local is None or rule.pattern is None:
                return None
            # Convert XPath regex to Python regex (best effort)
            try:
                pattern = _convert_xpath_pattern(rule.pattern)
                return AttributeValuePatternConstraint(
                    attribute=attr_local,
                    pattern=pattern,
                    namespace=attr_namespace,
                )
            except Exception:
                # Pattern cannot be converted
                return None

        case SchematronType.UNIQUE_ATTRIBUTE:
            if attr_local is None:
                return None
            element_tag = _resolve_context_to_tag(rule.context, namespace_map)
            return UniqueAttributeValueConstraint(
                attribute=attr_local,
                namespace=attr_namespace,
                element_tag=element_tag,
            )

        case SchematronType.RELATIONSHIP_TYPE:
            if attr_local is None or rule.relationship_type is None:
                return None
            return RelationshipTypeConstraint(
                relationship_id_attribute=attr_local,
                expected_type=rule.relationship_type,
                namespace=attr_namespace,
            )

        case SchematronType.ELEMENT_REFERENCE:
            # Element references need more context - skip for now
            return None

        case SchematronType.ATTRIBUTE_NOT_EQUAL:
            if attr_local is None or rule.forbidden_value is None:
                return None
            return AttributeNotEqualConstraint(
                attribute=attr_local,
                forbidden_value=rule.forbidden_value,
                namespace=attr_namespace,
            )

        case SchematronType.ATTRIBUTE_EQUALS:
            if attr_local is None or rule.expected_value is None:
                return None
            return AttributeEqualsConstraint(
                attribute=attr_local,
                expected_value=rule.expected_value,
                namespace=attr_namespace,
            )

        case SchematronType.ATTRIBUTE_COMPARISON:
            if (
                attr_local is None
                or rule.other_attribute is None
                or rule.comparison_operator is None
            ):
                return None
            other_local, other_namespace = _split_attribute_name(
                rule.other_attribute, namespace_map
            )
            namespace = attr_namespace if attr_namespace == other_namespace else None
            return AttributeComparisonConstraint(
                attribute=attr_local,
                other_attribute=other_local,
                operator=rule.comparison_operator,
                namespace=namespace,
            )

        case SchematronType.OR_CONDITION:
            if not rule.sub_rules:
                return None
            sub_constraints = []
            for sub_rule in rule.sub_rules:
                sub_constraint = create_constraint_from_schematron(
                    sub_rule, namespace_map
                )
                if sub_constraint is not None:
                    sub_constraints.append(sub_constraint)
            if not sub_constraints:
                return None
            return OrConstraint(constraints=sub_constraints)

        case SchematronType.AND_CONDITION:
            if not rule.sub_rules:
                return None
            sub_constraints = []
            for sub_rule in rule.sub_rules:
                sub_constraint = create_constraint_from_schematron(
                    sub_rule, namespace_map
                )
                if sub_constraint is not None:
                    sub_constraints.append(sub_constraint)
            if not sub_constraints:
                return None
            return AndConstraint(constraints=sub_constraints)

        case SchematronType.ATTRIBUTES_PRESENT:
            if not rule.required_attributes:
                return None
            attr_items = [_split_attribute_name(a, namespace_map) for a in rule.required_attributes]
            attr_names = [local for local, _ in attr_items]
            namespaces = {ns for _, ns in attr_items if ns}
            namespace = namespaces.pop() if len(namespaces) == 1 else None
            return AttributesPresentConstraint(
                attributes=attr_names,
                namespace=namespace,
                all_required=True,
            )

        case SchematronType.CROSS_PART_COUNT:
            if (
                attr_local is None
                or rule.part_path is None
                or rule.element_xpath is None
            ):
                return None
            return CrossPartCountConstraint(
                attribute=attr_local,
                attribute_namespace=attr_namespace,
                part_path=rule.part_path,
                element_xpath=rule.element_xpath,
                count_offset=rule.count_offset,
                namespace_map=namespace_map,
            )

        case SchematronType.CONDITIONAL_VALUE:
            # @attr and <condition> -> if attr present, condition must hold
            if attr_local is None or not rule.sub_rules:
                return None
            sub_constraint = create_constraint_from_schematron(
                rule.sub_rules[0], namespace_map
            )
            if sub_constraint is None:
                return None
            return ConditionalConstraint(
                trigger_attribute=attr_local,
                constraint=sub_constraint,
                namespace=attr_namespace,
            )

        case SchematronType.UNKNOWN:
            return None

    return None


def load_sdk_constraints(
    app_filter: str = "All",
) -> Iterator[tuple[str, SemanticConstraint]]:
    """Load all interpretable SDK schematron rules as constraints.

    Args:
        app_filter: Filter by application ("All", "PowerPoint", "Word", "Excel")

    Yields:
        Tuples of (element_tag, constraint) for each convertible rule.
    """
    namespace_map = _get_namespace_map()
    schematron_registry = get_schematron_registry()
    schematron_registry.load()

    for rule in schematron_registry._rules:
        # Filter by application
        if app_filter != "All" and rule.app not in ("All", app_filter):
            continue

        # Convert context to Clark notation tag
        element_tag = _resolve_context_to_tag(rule.context, namespace_map)
        if element_tag is None:
            continue

        # Convert rule to constraint
        constraint = create_constraint_from_schematron(rule, namespace_map)
        if constraint is None:
            continue

        yield element_tag, constraint


def get_sdk_constraint_stats() -> dict:
    """Get statistics about SDK constraint conversion.

    Returns:
        Dictionary with conversion statistics.
    """
    namespace_map = _get_namespace_map()
    schematron_registry = get_schematron_registry()
    schematron_registry.load()

    stats = {
        "total": 0,
        "converted": 0,
        "skipped_no_context": 0,
        "skipped_no_constraint": 0,
        "by_type": {},
    }

    for rule in schematron_registry._rules:
        stats["total"] += 1

        # Track by type
        type_name = rule.rule_type.name
        if type_name not in stats["by_type"]:
            stats["by_type"][type_name] = {"total": 0, "converted": 0}
        stats["by_type"][type_name]["total"] += 1

        # Try to convert
        element_tag = _resolve_context_to_tag(rule.context, namespace_map)
        if element_tag is None:
            stats["skipped_no_context"] += 1
            continue

        constraint = create_constraint_from_schematron(rule, namespace_map)
        if constraint is None:
            stats["skipped_no_constraint"] += 1
            continue

        stats["converted"] += 1
        stats["by_type"][type_name]["converted"] += 1

    return stats
