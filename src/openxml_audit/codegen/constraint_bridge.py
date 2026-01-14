"""Bridge between SDK schema data and runtime constraint validation.

Converts SDK types to our constraint classes on-demand.
"""

from __future__ import annotations

from functools import lru_cache
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.codegen.schema_loader import (
    SdkAttribute,
    SdkElementType,
    SdkParticle,
    get_registry,
    get_xsd_type_name,
)
from openxml_audit.schema.constraints import (
    AttributeConstraint,
    ElementConstraint,
)
from openxml_audit.schema.particle import (
    AllParticle,
    AnyParticle,
    ChoiceParticle,
    ElementParticle,
    CompositeParticle,
    ParticleConstraint,
    SequenceParticle,
)
from openxml_audit.schema.types import (
    XsdTypeValidator,
    get_type_validator,
)

if TYPE_CHECKING:
    pass


def _convert_attribute(
    attr: SdkAttribute,
    namespace_map: dict[str, str],
) -> AttributeConstraint:
    """Convert an SDK attribute to an AttributeConstraint."""
    # Determine namespace from prefix
    ns = None
    if attr.prefix:
        ns = namespace_map.get(attr.prefix)

    # Get type validator
    xsd_type = get_xsd_type_name(attr.type_name)
    type_validator = get_type_validator(xsd_type)

    return AttributeConstraint(
        namespace=ns,
        local_name=attr.local_name,
        type_validator=type_validator,
        required=attr.required,
    )


def _convert_particle(
    particle: SdkParticle,
    target_namespace: str,
    namespace_map: dict[str, str],
) -> ParticleConstraint | None:
    """Convert an SDK particle to a ParticleConstraint."""
    kind = particle.kind

    def maybe_collapse_single_child(
        children: list[ParticleConstraint],
    ) -> ParticleConstraint | None:
        if len(children) != 1:
            return None
        if particle.min_occurs != 1 or particle.max_occurs != 1:
            return None
        return children[0]

    if kind == "Sequence":
        children = [
            _convert_particle(item, target_namespace, namespace_map)
            for item in particle.items
        ]
        children = [c for c in children if c is not None]
        flattened: list[ParticleConstraint] = []
        for child in children:
            if isinstance(child, SequenceParticle) and child.max_occurs == 1:
                optional = child.min_occurs == 0
                if optional and any(sub.min_occurs > 0 for sub in child.children):
                    flattened.append(child)
                    continue
                for sub in child.children:
                    if optional and sub.min_occurs > 0:
                        sub.min_occurs = 0
                    flattened.append(sub)
            else:
                flattened.append(child)
        children = flattened
        collapsed = maybe_collapse_single_child(children)
        if collapsed is not None:
            return collapsed
        return SequenceParticle(
            children=children,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif kind == "Choice":
        children = [
            _convert_particle(item, target_namespace, namespace_map)
            for item in particle.items
        ]
        children = [c for c in children if c is not None]
        flattened = []
        for child in children:
            if (
                isinstance(child, ChoiceParticle)
                and child.min_occurs == 1
                and child.max_occurs == 1
            ):
                flattened.extend(child.children)
            else:
                flattened.append(child)
        children = flattened
        collapsed = maybe_collapse_single_child(children)
        if collapsed is not None:
            return collapsed
        return ChoiceParticle(
            children=children,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif kind == "All":
        children = [
            _convert_particle(item, target_namespace, namespace_map)
            for item in particle.items
        ]
        children = [c for c in children if c is not None]
        collapsed = maybe_collapse_single_child(children)
        if collapsed is not None:
            return collapsed
        return AllParticle(
            children=children,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif kind == "Group":
        # Group is a reference to a named group - inline its items
        children = [
            _convert_particle(item, target_namespace, namespace_map)
            for item in particle.items
        ]
        children = [c for c in children if c is not None]
        if len(children) == 1:
            # Single child group - just return the child
            child = children[0]
            # Apply group's occurrence to child
            if isinstance(child, (SequenceParticle, ChoiceParticle, AllParticle)):
                child.min_occurs = particle.min_occurs
                child.max_occurs = particle.max_occurs
            return child
        else:
            # Multiple children in group - wrap in sequence
            return SequenceParticle(
                children=children,
                min_occurs=particle.min_occurs,
                max_occurs=particle.max_occurs,
            )

    elif kind == "Any":
        return AnyParticle(
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif particle.name:
        # Element reference like "a:CT_OfficeArtExtensionList/a:extLst"
        # Extract the element local name
        if "/" in particle.name:
            elem_ref = particle.name.split("/")[1]
            if ":" in elem_ref:
                prefix, local_name = elem_ref.split(":", 1)
                ns = namespace_map.get(prefix, target_namespace)
            else:
                local_name = elem_ref
                ns = target_namespace
        else:
            local_name = particle.name.split(":")[-1]
            ns = target_namespace

        return ElementParticle(
            namespace=ns,
            local_name=local_name,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    return None


def _build_namespace_map() -> dict[str, str]:
    """Build prefix -> namespace URI map from registry."""
    registry = get_registry()
    registry.load()
    return dict(registry._prefixes)


@lru_cache(maxsize=2048)
def get_element_constraint(tag: str) -> ElementConstraint | None:
    """Get element constraint from SDK schema by Clark notation tag.

    Args:
        tag: Element tag in Clark notation, e.g., "{namespace}localname"

    Returns:
        ElementConstraint if found, None otherwise.
    """
    registry = get_registry()
    elem_type = registry.get_element_type_by_tag(tag)

    if elem_type is None:
        return None

    return convert_element_type(elem_type)


def get_element_constraint_for_element(
    tag: str, element: etree._Element
) -> ElementConstraint | None:
    """Get the best element constraint for a specific element instance."""
    registry = get_registry()
    candidates = registry.get_element_type_candidates(tag)

    if not candidates:
        elem_type = registry.get_element_type_by_tag(tag)
        return convert_element_type(elem_type) if elem_type else None

    if len(candidates) == 1:
        return convert_element_type(candidates[0])

    children = [c for c in element if isinstance(c.tag, str)]
    best: ElementConstraint | None = None
    best_score: tuple[int, int] | None = None

    for candidate in candidates:
        constraint = convert_element_type(candidate)
        if _missing_required_attributes(constraint, element):
            continue
        score = _score_candidate(constraint, children)
        if best_score is None or score > best_score:
            best = constraint
            best_score = score

    if best is not None:
        return best

    elem_type = registry.get_element_type_by_tag(tag)
    return convert_element_type(elem_type) if elem_type else None


def convert_element_type(elem_type: SdkElementType) -> ElementConstraint:
    """Convert an SDK element type to an ElementConstraint.

    Args:
        elem_type: The SDK element type to convert.

    Returns:
        The converted ElementConstraint.
    """
    registry = get_registry()
    namespace_map = _build_namespace_map()

    # Get schema for target namespace
    schema = None
    for ns, s in registry._schemas.items():
        if elem_type in s.types:
            schema = s
            break

    target_namespace = schema.target_namespace if schema else ""

    # Convert attributes
    attributes = [
        _convert_attribute(attr, namespace_map)
        for attr in elem_type.attributes
    ]

    # Convert particle (content model)
    content_model = None
    if elem_type.particle:
        content_model = _convert_particle(
            elem_type.particle,
            target_namespace,
            namespace_map,
        )

    # Determine namespace from element name
    if elem_type.element_prefix:
        ns = namespace_map.get(elem_type.element_prefix, target_namespace)
    else:
        ns = target_namespace

    return ElementConstraint(
        namespace=ns,
        local_name=elem_type.element_name or "",
        attributes=attributes,
        content_model=content_model,
        allows_text=False,  # Could be determined from base class
    )


def _missing_required_attributes(
    constraint: ElementConstraint,
    element: etree._Element,
) -> bool:
    for attr in constraint.get_required_attributes():
        if attr.qualified_name not in element.attrib:
            return True
    return False


def _score_candidate(
    constraint: ElementConstraint,
    children: list[etree._Element],
) -> tuple[int, int]:
    if constraint.content_model is None or not children:
        return (0, 0)

    allowed, has_any = _collect_allowed_tags(constraint.content_model)
    if has_any:
        total_matches = len(children)
    else:
        total_matches = sum(1 for child in children if child.tag in allowed)
    specific_matches = sum(1 for child in children if child.tag in allowed)
    return (specific_matches, total_matches)


def _collect_allowed_tags(
    particle: ParticleConstraint,
) -> tuple[set[str], bool]:
    allowed: set[str] = set()
    has_any = False

    def visit(node: ParticleConstraint) -> None:
        nonlocal has_any
        if isinstance(node, ElementParticle):
            allowed.add(node.qualified_name)
        elif isinstance(node, AnyParticle):
            has_any = True
        elif isinstance(node, CompositeParticle):
            for child in node.children:
                visit(child)

    visit(particle)
    return allowed, has_any


def get_sdk_element_info(tag: str) -> dict | None:
    """Get raw SDK element info for debugging/inspection.

    Args:
        tag: Element tag in Clark notation.

    Returns:
        Dictionary with element info, or None.
    """
    registry = get_registry()
    elem_type = registry.get_element_type_by_tag(tag)

    if elem_type is None:
        return None

    return {
        "name": elem_type.name,
        "class_name": elem_type.class_name,
        "base_class": elem_type.base_class,
        "is_abstract": elem_type.is_abstract,
        "is_leaf": elem_type.is_leaf_element,
        "attributes": [
            {
                "qname": a.qname,
                "type": a.type_name,
                "required": a.required,
            }
            for a in elem_type.attributes
        ],
        "has_particle": elem_type.particle is not None,
    }
