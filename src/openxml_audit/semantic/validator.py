"""Semantic validator for Open XML documents."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.context import ElementContext
from openxml_audit.errors import FileFormat, ValidationError
from openxml_audit.namespaces import MC, OFFICE_DOC_RELATIONSHIPS
from openxml_audit.semantic.attributes import SemanticConstraint
from openxml_audit.semantic.references import IdTracker, validate_unique_ids
from openxml_audit.semantic.relationships import validate_part_relationships

_MC_ALTERNATE_CONTENT = f"{{{MC}}}AlternateContent"
_MC_NS = {"mc": MC}

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext
    from openxml_audit.parts import OpenXmlPart


class SemanticValidator:
    """Validates XML documents against semantic constraints.

    This validator checks:
    - Attribute value relationships and dependencies
    - Relationship validity and targets
    - ID uniqueness and references
    - Cross-element constraints
    """

    def __init__(self, validate_unique_ids: bool = True) -> None:
        self._constraints: dict[str, list[SemanticConstraint]] = {}
        self._id_tracker = IdTracker()
        self._validate_unique_ids = validate_unique_ids
        self._known_namespaces: set[str] | None = None
        self._namespace_versions: dict[str, FileFormat] | None = None

    def register_constraint(self, element_tag: str, constraint: SemanticConstraint) -> None:
        """Register a semantic constraint for an element type.

        Args:
            element_tag: The element tag (Clark notation) to apply constraint to.
            constraint: The constraint to apply.
        """
        if element_tag not in self._constraints:
            self._constraints[element_tag] = []
        self._constraints[element_tag].append(constraint)

    def validate_part(
        self, part: OpenXmlPart, context: ValidationContext
    ) -> list[ValidationError]:
        """Validate a part against semantic constraints.

        Args:
            part: The part to validate.
            context: The validation context.

        Returns:
            List of validation errors.
        """
        context.set_part(part)

        # Validate relationships
        validate_part_relationships(part, context)

        xml = part.xml
        if xml is None:
            return context.errors

        # Clear ID tracker for this part
        self._id_tracker.clear(part.uri)

        # Validate unique IDs when enabled
        if self._validate_unique_ids:
            validate_unique_ids(xml, "id", context, self._id_tracker, part.uri)

        # Validate element constraints
        self._validate_element(xml, context)

        return context.errors

    def _validate_element(self, element: etree._Element, context: ValidationContext) -> None:
        """Validate an element and its children recursively."""
        if context.should_stop:
            return

        with ElementContext(context, element):
            tag = element.tag

            # Validate relationship attributes for any OOXML element.
            self._validate_relationship_attributes(element, context)
            self._validate_mc_ignorable(element, context)

            # Apply registered constraints for this element type
            if tag in self._constraints:
                for constraint in self._constraints[tag]:
                    constraint.validate(element, context)

            # Recursively validate children.
            # At Office2010, the SDK resolves mc:AlternateContent before
            # semantic validation (skipping mc:Fallback when mc:Choice is
            # understood).  At all other versions, both branches are
            # validated.
            mce_resolve = context.file_format == FileFormat.OFFICE_2010
            for child in element:
                if not isinstance(child.tag, str):
                    continue
                if mce_resolve and child.tag == _MC_ALTERNATE_CONTENT:
                    for resolved in self._resolve_mce(child, context.file_format):
                        self._validate_element(resolved, context)
                else:
                    self._validate_element(child, context)

    def _validate_relationship_attributes(
        self, element: etree._Element, context: ValidationContext
    ) -> None:
        """Ensure relationship ID attributes reference existing relationships."""
        part = context.part
        if part is None:
            return
        for attr_name, value in element.attrib.items():
            if not attr_name.startswith(f"{{{OFFICE_DOC_RELATIONSHIPS}}}"):
                continue
            if not value:
                continue
            rel = part.relationships.get_by_id(value)
            if rel is None:
                local_attr = attr_name.split("}")[-1] if attr_name.startswith("{") else attr_name
                context.add_semantic_error(
                    f"The relationship '{value}' referenced by attribute "
                    f"'{local_attr}' does not exist.",
                    node=local_attr,
                    error_id="Sem_MissingRelationshipReference",
                )

    def _validate_mc_ignorable(
        self, element: etree._Element, context: ValidationContext
    ) -> None:
        ignorable = element.get(f"{{{MC}}}Ignorable")
        if ignorable is None:
            return
        prefixes = [prefix for prefix in ignorable.split() if prefix]
        if not prefixes:
            return
        nsmap = element.nsmap or {}
        for prefix in prefixes:
            if prefix not in nsmap or not nsmap.get(prefix):
                context.add_semantic_error(
                    f"Ignorable attribute contains undefined prefix '{prefix}'",
                    node="Ignorable",
                )

    # ------------------------------------------------------------------
    # MCE resolution (Office 2010 only)
    # ------------------------------------------------------------------

    def _ensure_namespace_data(self) -> None:
        """Lazily load known namespaces and version map from the schema registry."""
        if self._known_namespaces is not None:
            return
        from openxml_audit.codegen.schema_loader import get_registry as get_schema_registry

        registry = get_schema_registry()
        if not registry._schemas:
            registry.load()

        known: set[str] = {MC}
        known.update(registry._prefixes.values())
        self._known_namespaces = known

        ns_versions: dict[str, FileFormat] = {}
        for ns, schema in registry._schemas.items():
            min_fmt: FileFormat | None = None
            min_order = float("inf")
            has_base = False
            for t in schema.types:
                fmt = FileFormat.from_version_string(t.version) if t.version else None
                if fmt is None:
                    has_base = True
                    break
                order = fmt.ooxml_order()
                if order < min_order:
                    min_order = order
                    min_fmt = fmt
            if not has_base and min_fmt is not None:
                ns_versions[ns] = min_fmt
        self._namespace_versions = ns_versions

    def _resolve_mce(
        self, alt: etree._Element, file_format: FileFormat
    ) -> list[etree._Element]:
        """Resolve mc:AlternateContent to the appropriate branch.

        At Office2010, the SDK resolves MCE: if a Choice's required namespaces
        are understood at the target version, its children are used; otherwise
        Fallback children are used.
        """
        self._ensure_namespace_data()
        assert self._known_namespaces is not None
        assert self._namespace_versions is not None

        # Try each Choice
        for choice in alt.findall("mc:Choice", _MC_NS):
            requires = choice.get("Requires", "").split()
            if not requires:
                return [c for c in choice if isinstance(c.tag, str)]
            if self._choice_is_understood(choice, requires, file_format):
                return [c for c in choice if isinstance(c.tag, str)]

        # Fall back
        fallback = alt.find("mc:Fallback", _MC_NS)
        if fallback is not None:
            return [c for c in fallback if isinstance(c.tag, str)]
        return []

    def _choice_is_understood(
        self,
        choice: etree._Element,
        required_prefixes: list[str],
        file_format: FileFormat,
    ) -> bool:
        assert self._known_namespaces is not None
        assert self._namespace_versions is not None
        nsmap = choice.nsmap or {}
        for prefix in required_prefixes:
            namespace = nsmap.get(prefix)
            if namespace is None:
                return False
            if namespace not in self._known_namespaces:
                return False
            ns_version = self._namespace_versions.get(namespace)
            if ns_version is not None and not file_format.includes_ooxml(ns_version):
                return False
        return True


def create_pptx_semantic_validator(load_sdk_rules: bool = True) -> SemanticValidator:
    """Create a semantic validator with PPTX-specific constraints.

    Args:
        load_sdk_rules: If True, load all SDK schematron rules.

    Returns:
        A SemanticValidator configured for PPTX validation.
    """
    from openxml_audit.namespaces import PRESENTATIONML
    from openxml_audit.semantic.relationships import RelationshipExistConstraint

    # SDK schematron rules already cover PPT ID-uniqueness semantics with
    # narrower scopes. The generic @id scan is too broad and causes false
    # positives on valid presentations.
    validator = SemanticValidator(validate_unique_ids=False)

    # Load SDK schematron rules
    if load_sdk_rules:
        try:
            from openxml_audit.codegen.schematron_bridge import load_sdk_constraints

            for element_tag, constraint in load_sdk_constraints(app_filter="All"):
                validator.register_constraint(element_tag, constraint)
        except ImportError:
            pass  # SDK data not available

    # Relationship namespace
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    # Slide ID references (keep these as they check relationship existence, not just type)
    validator.register_constraint(
        f"{{{PRESENTATIONML}}}sldId",
        RelationshipExistConstraint(
            attribute="id",
            namespace=rel_ns,
        ),
    )

    # Slide master ID references
    validator.register_constraint(
        f"{{{PRESENTATIONML}}}sldMasterId",
        RelationshipExistConstraint(
            attribute="id",
            namespace=rel_ns,
        ),
    )

    # Notes master ID references
    validator.register_constraint(
        f"{{{PRESENTATIONML}}}notesMasterId",
        RelationshipExistConstraint(
            attribute="id",
            namespace=rel_ns,
        ),
    )

    return validator


def create_word_semantic_validator(load_sdk_rules: bool = True) -> SemanticValidator:
    """Create a semantic validator with Word-specific constraints."""
    from openxml_audit.namespaces import OFFICE_DOC_RELATIONSHIPS, REL_FONT, WORDPROCESSINGML
    from openxml_audit.semantic.attributes import (
        AttributeValueInSetConstraint,
        AttributeValuePatternConstraint,
    )
    from openxml_audit.semantic.relationships import RelationshipTypeConstraint

    validator = SemanticValidator(validate_unique_ids=False)

    if load_sdk_rules:
        try:
            from openxml_audit.codegen.schematron_bridge import load_sdk_constraints

            for element_tag, constraint in load_sdk_constraints(app_filter="Word"):
                validator.register_constraint(element_tag, constraint)
        except ImportError:
            pass

    zoom_values = ("none", "fullPage", "bestFit", "textFit")
    validator.register_constraint(
        f"{{{WORDPROCESSINGML}}}zoom",
        AttributeValueInSetConstraint(
            attribute="val",
            namespace=WORDPROCESSINGML,
            allowed_values=zoom_values,
        ),
    )

    font_key_pattern = (
        r"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
        r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$"
    )
    for tag in ("embedRegular", "embedBold", "embedItalic", "embedBoldItalic"):
        element_tag = f"{{{WORDPROCESSINGML}}}{tag}"
        validator.register_constraint(
            element_tag,
            AttributeValuePatternConstraint(
                attribute="fontKey",
                namespace=WORDPROCESSINGML,
                pattern=font_key_pattern,
            ),
        )
        validator.register_constraint(
            element_tag,
            RelationshipTypeConstraint(
                relationship_id_attribute="id",
                namespace=OFFICE_DOC_RELATIONSHIPS,
                expected_type=REL_FONT,
            ),
        )

    return validator


def create_spreadsheet_semantic_validator(load_sdk_rules: bool = True) -> SemanticValidator:
    """Create a semantic validator with Excel-specific constraints."""
    validator = SemanticValidator(validate_unique_ids=False)

    if load_sdk_rules:
        try:
            from openxml_audit.codegen.schematron_bridge import load_sdk_constraints

            for element_tag, constraint in load_sdk_constraints(app_filter="Excel"):
                validator.register_constraint(element_tag, constraint)
        except ImportError:
            pass

    return validator
