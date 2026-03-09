"""Schema validator for Open XML documents."""

from __future__ import annotations

from time import perf_counter
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.codegen.schema_loader import get_registry as get_schema_registry
from openxml_audit.context import ElementContext, ValidationContext
from openxml_audit.errors import FileFormat, ValidationError
from openxml_audit.namespaces import MC, WORDPROCESSINGML
from openxml_audit.schema.constraints import get_constraint_for_tag as get_hardcoded_constraint
from openxml_audit.schema.particle import (
    CompositeParticle,
    get_validator,
)

# Try to import SDK constraint bridge
try:
    from openxml_audit.codegen.constraint_bridge import (
        get_element_constraint as get_sdk_constraint,
    )
    from openxml_audit.codegen.constraint_bridge import (
        get_element_constraint_for_element as get_sdk_constraint_for_element,
    )
    _HAS_SDK_CONSTRAINTS = True
except ImportError:
    _HAS_SDK_CONSTRAINTS = False
    get_sdk_constraint = None  # type: ignore
    get_sdk_constraint_for_element = None  # type: ignore


def get_constraint_for_tag(tag: str, element: etree._Element | None = None):
    """Get constraint for element tag, preferring SDK constraints.

    Args:
        tag: Element tag in Clark notation.

    Returns:
        ElementConstraint if found, None otherwise.
    """
    # Try SDK constraints first (more complete)
    if _HAS_SDK_CONSTRAINTS and get_sdk_constraint is not None:
        if element is not None and get_sdk_constraint_for_element is not None:
            constraint = get_sdk_constraint_for_element(tag, element)
        else:
            constraint = get_sdk_constraint(tag)
        if constraint is not None:
            return constraint

    # Fall back to hardcoded constraints
    return get_hardcoded_constraint(tag)

if TYPE_CHECKING:
    from openxml_audit.parts import OpenXmlPart


class SchemaValidator:
    """Validates XML documents against schema constraints.

    This validator checks:
    - Required attributes are present
    - Attribute values conform to type constraints
    - Child elements match content model (sequence, choice, all)
    """

    def __init__(self, validate_unknown_elements: bool = False):
        """Initialize the schema validator.

        Args:
            validate_unknown_elements: If True, report errors for elements
                                       without known constraints.
        """
        self._validate_unknown = validate_unknown_elements
        self._schema_registry = get_schema_registry()
        self._schema_registry.load()
        self._known_namespaces = self._build_known_namespaces()
        self._undeclared_validation_cache: dict[str, bool] = {}
        self._collect_metrics = False
        self._metrics: dict[str, float] = self._new_metrics()

    def _new_metrics(self) -> dict[str, float]:
        return {
            "elements": 0.0,
            "constraint_lookup": 0.0,
            "children_expand": 0.0,
            "attributes": 0.0,
            "content_model": 0.0,
            "recursion": 0.0,
        }

    def set_metric_collection(self, enabled: bool) -> None:
        """Enable/disable metric collection for schema hot paths."""
        self._collect_metrics = enabled

    def reset_metrics(self) -> None:
        """Reset collected schema hot-path metrics."""
        self._metrics = self._new_metrics()

    def get_metrics(self) -> dict[str, float]:
        """Get collected schema hot-path metrics."""
        return dict(self._metrics)

    def validate_part(
        self, part: OpenXmlPart, context: ValidationContext
    ) -> list[ValidationError]:
        """Validate an XML part against schema constraints.

        Args:
            part: The part to validate.
            context: The validation context.

        Returns:
            List of validation errors.
        """
        context.set_part(part)

        xml = part.xml
        if xml is None:
            return context.errors

        if self._collect_metrics:
            self._validate_element_with_metrics(xml, context)
        else:
            self._validate_element(xml, context)

        return context.errors

    def _validate_element(self, element: etree._Element, context: ValidationContext) -> None:
        """Validate an element and its children recursively."""
        if context.should_stop:
            return

        with ElementContext(context, element):
            tag = element.tag

            # Get constraint for this element
            constraint = get_constraint_for_tag(tag, element)

            children = self._get_validation_children(element, context)

            if constraint is not None:
                # Validate attributes
                self._validate_attributes(element, constraint, context)

                # Validate content model
                if constraint.content_model is not None:
                    self._validate_content_model(constraint.content_model, children, context)

            # Recursively validate children
            for child in children:
                self._validate_element(child, context)

    def _validate_element_with_metrics(
        self, element: etree._Element, context: ValidationContext
    ) -> None:
        """Validate an element recursively while collecting hot-path timings."""
        if context.should_stop:
            return

        with ElementContext(context, element):
            self._metrics["elements"] += 1.0
            tag = element.tag

            lookup_start = perf_counter()
            constraint = get_constraint_for_tag(tag, element)
            self._metrics["constraint_lookup"] += perf_counter() - lookup_start

            children_start = perf_counter()
            children = self._get_validation_children(element, context)
            self._metrics["children_expand"] += perf_counter() - children_start

            if constraint is not None:
                attr_start = perf_counter()
                self._validate_attributes(element, constraint, context)
                self._metrics["attributes"] += perf_counter() - attr_start

                if constraint.content_model is not None:
                    model_start = perf_counter()
                    self._validate_content_model(constraint.content_model, children, context)
                    self._metrics["content_model"] += perf_counter() - model_start

            recurse_start = perf_counter()
            for child in children:
                self._validate_element_with_metrics(child, context)
            self._metrics["recursion"] += perf_counter() - recurse_start

    def _validate_attributes(
        self,
        element: etree._Element,
        constraint: ElementConstraint,  # type: ignore
        context: ValidationContext,
    ) -> None:
        """Validate element attributes."""
        # Check required attributes
        for attr_constraint in self._get_required_attributes(constraint):
            attr_name = attr_constraint.qualified_name
            if attr_name not in element.attrib:
                context.add_schema_error(
                    f"Required attribute '{attr_constraint.local_name}' is missing",
                    node=attr_constraint.local_name,
                )

        # Validate attribute values
        for attr_constraint in constraint.attributes:
            attr_name = attr_constraint.qualified_name
            if attr_name in element.attrib:
                value = element.attrib[attr_name]

                # Check fixed value
                if (
                    attr_constraint.fixed_value is not None
                    and value != attr_constraint.fixed_value
                ):
                    context.add_schema_error(
                        f"Attribute '{attr_constraint.local_name}' must have "
                        f"fixed value '{attr_constraint.fixed_value}', got '{value}'",
                        node=attr_constraint.local_name,
                    )

                # Type validation
                if attr_constraint.type_validator is not None:
                    result = attr_constraint.type_validator.validate(value, context)
                    if not result.is_valid:
                        context.add_schema_error(
                            f"Invalid value for attribute '{attr_constraint.local_name}': "
                            f"{result.error_message}",
                            node=attr_constraint.local_name,
                        )

        # Check undeclared attributes only when SDK metadata is unambiguous.
        if not self._should_validate_undeclared_attributes(element, constraint):
            return

        declared = self._get_declared_attributes(constraint)
        element_ns = self._extract_namespace(element.tag)
        for attr_name in element.attrib:
            if attr_name in declared:
                continue

            attr_ns = self._extract_namespace(attr_name)
            attr_local = self._extract_local_name(attr_name)

            # Markup compatibility and xml:* attributes are handled separately.
            if attr_ns == MC:
                continue
            if attr_ns == "http://www.w3.org/XML/1998/namespace":
                continue

            # Foreign-namespace extension attributes are generally allowed.
            if attr_ns is not None and attr_ns != element_ns:
                continue

            context.add_schema_error(
                f"The '{attr_local}' attribute is not declared.",
                node=attr_local,
            )


    def _validate_content_model(
        self,
        content_model: ParticleConstraint,  # type: ignore
        children: list[etree._Element],
        context: ValidationContext,
    ) -> None:
        """Validate element children against content model."""
        if isinstance(content_model, CompositeParticle):
            validator = get_validator(content_model.particle_type)
            if validator is not None:
                validator.validate(content_model, children, context)

    def _get_validation_children(
        self, element: etree._Element, context: ValidationContext
    ) -> list[etree._Element]:
        children: list[etree._Element] = []
        ignorable_namespaces = self._collect_ignorable_namespaces(element)

        for child in element:
            if not isinstance(child.tag, str):
                continue
            child_ns = self._extract_namespace(child.tag)
            if child_ns is not None and child_ns in ignorable_namespaces:
                continue
            if self._is_version_ignored_child(element.tag, child.tag, context.file_format):
                continue
            if child.tag == f"{{{MC}}}AlternateContent":
                children.extend(self._resolve_alternate_content(child))
                continue
            children.append(child)

        return children

    def _is_version_ignored_child(
        self, parent_tag: str, child_tag: str, file_format: FileFormat
    ) -> bool:
        if (
            parent_tag == f"{{{WORDPROCESSINGML}}}settings"
            and child_tag == f"{{{WORDPROCESSINGML}}}doNotEmbedSmartTags"
        ):
            return file_format in {
                FileFormat.OFFICE_2013,
                FileFormat.OFFICE_2016,
                FileFormat.OFFICE_2019,
                FileFormat.OFFICE_2021,
                FileFormat.MICROSOFT_365,
            }
        return False

    def _collect_ignorable_namespaces(self, element: etree._Element) -> set[str]:
        namespaces: set[str] = set()
        current: etree._Element | None = element
        while current is not None and isinstance(current.tag, str):
            ignorable = current.get(f"{{{MC}}}Ignorable")
            if ignorable:
                nsmap = current.nsmap or {}
                for prefix in ignorable.split():
                    namespace = nsmap.get(prefix)
                    if namespace:
                        namespaces.add(namespace)
            parent = current.getparent()
            current = parent if isinstance(parent, etree._Element) else None
        return namespaces

    def _resolve_alternate_content(self, alt: etree._Element) -> list[etree._Element]:
        ns = {"mc": MC}
        chosen = self._select_alternate_content_choice(alt, ns)
        if chosen is None:
            chosen = alt.find("mc:Fallback", ns)
        if chosen is None:
            return []
        return [c for c in chosen if isinstance(c.tag, str)]

    def _select_alternate_content_choice(
        self, alt: etree._Element, namespaces: dict[str, str]
    ) -> etree._Element | None:
        for choice in alt.findall("mc:Choice", namespaces):
            requires = choice.get("Requires", "").split()
            if not requires:
                return choice
            if self._choice_is_understood(choice, requires):
                return choice
        return None

    def _choice_is_understood(
        self, choice: etree._Element, required_prefixes: list[str]
    ) -> bool:
        nsmap = choice.nsmap or {}
        for prefix in required_prefixes:
            namespace = nsmap.get(prefix)
            if namespace is None:
                return False
            if namespace not in self._known_namespaces:
                return False
        return True

    def _build_known_namespaces(self) -> set[str]:
        known = {MC}
        known.update(self._schema_registry._prefixes.values())
        return known

    def _extract_namespace(self, qname: str) -> str | None:
        if qname.startswith("{") and "}" in qname:
            return qname[1:].split("}", 1)[0]
        return None

    def _extract_local_name(self, qname: str) -> str:
        if qname.startswith("{") and "}" in qname:
            return qname.split("}", 1)[1]
        return qname

    def _should_validate_undeclared_attributes(
        self,
        element: etree._Element,
        constraint: ElementConstraint,
    ) -> bool:
        # SDK metadata currently omits inherited attributes for a subset of elements.
        # Restrict undeclared checks to elements with explicit attribute declarations.
        if not constraint.attributes:
            return False

        tag = element.tag
        if tag not in self._undeclared_validation_cache:
            candidates = self._schema_registry.get_element_type_candidates(tag)
            self._undeclared_validation_cache[tag] = len(candidates) == 1
        return self._undeclared_validation_cache[tag]

    def _get_required_attributes(
        self, constraint: ElementConstraint
    ) -> tuple[AttributeConstraint, ...]:
        cached = getattr(constraint, "_oa_required_attrs_cache", None)
        if cached is None:
            cached = tuple(constraint.get_required_attributes())
            constraint._oa_required_attrs_cache = cached
        return cached

    def _get_declared_attributes(self, constraint: ElementConstraint) -> frozenset[str]:
        cached = getattr(constraint, "_oa_declared_attrs_cache", None)
        if cached is None:
            cached = frozenset(attr.qualified_name for attr in constraint.attributes)
            constraint._oa_declared_attrs_cache = cached
        return cached


# Import for type hints only
from openxml_audit.schema.constraints import AttributeConstraint, ElementConstraint  # noqa: E402
from openxml_audit.schema.particle import ParticleConstraint  # noqa: E402
