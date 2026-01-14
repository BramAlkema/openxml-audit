"""Element and attribute constraints for Open XML schema validation.

Defines the structure requirements for PPTX elements based on ECMA-376.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from openxml_audit.namespaces import DRAWINGML, PRESENTATIONML
from openxml_audit.schema.particle import (
    AnyParticle,
    ChoiceParticle,
    ElementParticle,
    ParticleConstraint,
    SequenceParticle,
)
from openxml_audit.schema.types import XsdBuiltinType, XsdTypeValidator, get_type_validator

if TYPE_CHECKING:
    pass


@dataclass
class AttributeConstraint:
    """Constraint for an XML attribute."""

    namespace: str | None
    local_name: str
    type_validator: XsdTypeValidator | None = None
    required: bool = False
    default_value: str | None = None
    fixed_value: str | None = None

    @property
    def qualified_name(self) -> str:
        """Get the Clark notation qualified name."""
        if self.namespace:
            return f"{{{self.namespace}}}{self.local_name}"
        return self.local_name


@dataclass
class ElementConstraint:
    """Constraint for an XML element type."""

    namespace: str
    local_name: str
    attributes: list[AttributeConstraint] = field(default_factory=list)
    content_model: ParticleConstraint | None = None
    allows_text: bool = False

    @property
    def qualified_name(self) -> str:
        """Get the Clark notation qualified name."""
        return f"{{{self.namespace}}}{self.local_name}"

    def get_attribute(self, name: str, namespace: str | None = None) -> AttributeConstraint | None:
        """Get an attribute constraint by name."""
        for attr in self.attributes:
            if attr.local_name == name and attr.namespace == namespace:
                return attr
        return None

    def get_required_attributes(self) -> list[AttributeConstraint]:
        """Get all required attributes."""
        return [attr for attr in self.attributes if attr.required]


class ElementConstraintRegistry:
    """Registry of element constraints for validation."""

    def __init__(self) -> None:
        self._constraints: dict[str, ElementConstraint] = {}

    def register(self, constraint: ElementConstraint) -> None:
        """Register an element constraint."""
        self._constraints[constraint.qualified_name] = constraint

    def get(self, qualified_name: str) -> ElementConstraint | None:
        """Get constraint for an element."""
        return self._constraints.get(qualified_name)

    def get_by_name(self, namespace: str, local_name: str) -> ElementConstraint | None:
        """Get constraint by namespace and local name."""
        return self.get(f"{{{namespace}}}{local_name}")


def _create_pptx_constraints() -> ElementConstraintRegistry:
    """Create constraints for core PPTX elements.

    This is a subset of the full ECMA-376 schema, covering the most
    critical elements for basic validation.
    """
    registry = ElementConstraintRegistry()

    # p:presentation - root element of presentation.xml
    presentation = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="presentation",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="saveSubsetFonts",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="autoCompressPictures",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
        ],
        content_model=SequenceParticle(children=[
            ElementParticle(PRESENTATIONML, "sldMasterIdLst", min_occurs=0),
            ElementParticle(PRESENTATIONML, "notesMasterIdLst", min_occurs=0),
            ElementParticle(PRESENTATIONML, "handoutMasterIdLst", min_occurs=0),
            ElementParticle(PRESENTATIONML, "sldIdLst", min_occurs=0),
            ElementParticle(PRESENTATIONML, "sldSz", min_occurs=0),
            ElementParticle(PRESENTATIONML, "notesSz", min_occurs=1),
            AnyParticle(namespace_constraint="##other", min_occurs=0, max_occurs=-1),
        ]),
    )
    registry.register(presentation)

    # p:sld - slide element
    slide = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="sld",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="showMasterSp",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="showMasterPhAnim",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="show",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
        ],
        content_model=SequenceParticle(children=[
            ElementParticle(PRESENTATIONML, "cSld", min_occurs=1),
            ElementParticle(PRESENTATIONML, "clrMapOvr", min_occurs=0),
            ElementParticle(PRESENTATIONML, "transition", min_occurs=0),
            ElementParticle(PRESENTATIONML, "timing", min_occurs=0),
            AnyParticle(namespace_constraint="##other", min_occurs=0, max_occurs=-1),
        ]),
    )
    registry.register(slide)

    # p:cSld - common slide data
    cSld = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="cSld",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="name",
                type_validator=get_type_validator(XsdBuiltinType.STRING),
                required=False,
            ),
        ],
        content_model=SequenceParticle(children=[
            ElementParticle(PRESENTATIONML, "bg", min_occurs=0),
            ElementParticle(PRESENTATIONML, "spTree", min_occurs=1),
            ElementParticle(PRESENTATIONML, "custDataLst", min_occurs=0),
            ElementParticle(PRESENTATIONML, "controls", min_occurs=0),
            AnyParticle(namespace_constraint="##other", min_occurs=0, max_occurs=-1),
        ]),
    )
    registry.register(cSld)

    # p:spTree - shape tree
    spTree = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="spTree",
        content_model=SequenceParticle(children=[
            ElementParticle(PRESENTATIONML, "nvGrpSpPr", min_occurs=1),
            ElementParticle(PRESENTATIONML, "grpSpPr", min_occurs=1),
            # Choice of shape types, unbounded
            ChoiceParticle(min_occurs=0, max_occurs=-1, children=[
                ElementParticle(PRESENTATIONML, "sp", min_occurs=1),
                ElementParticle(PRESENTATIONML, "grpSp", min_occurs=1),
                ElementParticle(PRESENTATIONML, "graphicFrame", min_occurs=1),
                ElementParticle(PRESENTATIONML, "cxnSp", min_occurs=1),
                ElementParticle(PRESENTATIONML, "pic", min_occurs=1),
                ElementParticle(PRESENTATIONML, "contentPart", min_occurs=1),
            ]),
            AnyParticle(namespace_constraint="##other", min_occurs=0, max_occurs=-1),
        ]),
    )
    registry.register(spTree)

    # p:sp - shape
    sp = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="sp",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="useBgFill",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
        ],
        content_model=SequenceParticle(children=[
            ElementParticle(PRESENTATIONML, "nvSpPr", min_occurs=1),
            ElementParticle(PRESENTATIONML, "spPr", min_occurs=1),
            ElementParticle(PRESENTATIONML, "style", min_occurs=0),
            ElementParticle(PRESENTATIONML, "txBody", min_occurs=0),
            AnyParticle(namespace_constraint="##other", min_occurs=0, max_occurs=-1),
        ]),
    )
    registry.register(sp)

    # p:nvSpPr - non-visual shape properties
    nvSpPr = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="nvSpPr",
        content_model=SequenceParticle(children=[
            ElementParticle(PRESENTATIONML, "cNvPr", min_occurs=1),
            ElementParticle(PRESENTATIONML, "cNvSpPr", min_occurs=1),
            ElementParticle(PRESENTATIONML, "nvPr", min_occurs=1),
        ]),
    )
    registry.register(nvSpPr)

    # p:cNvPr - common non-visual properties
    cNvPr = ElementConstraint(
        namespace=PRESENTATIONML,
        local_name="cNvPr",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="id",
                type_validator=get_type_validator(XsdBuiltinType.UNSIGNED_INT),
                required=True,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="name",
                type_validator=get_type_validator(XsdBuiltinType.STRING),
                required=True,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="descr",
                type_validator=get_type_validator(XsdBuiltinType.STRING),
                required=False,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="hidden",
                type_validator=get_type_validator(XsdBuiltinType.BOOLEAN),
                required=False,
            ),
        ],
    )
    registry.register(cNvPr)

    # a:off - offset in DrawingML
    off = ElementConstraint(
        namespace=DRAWINGML,
        local_name="off",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="x",
                type_validator=get_type_validator(XsdBuiltinType.LONG),
                required=True,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="y",
                type_validator=get_type_validator(XsdBuiltinType.LONG),
                required=True,
            ),
        ],
    )
    registry.register(off)

    # a:ext - extents in DrawingML
    ext = ElementConstraint(
        namespace=DRAWINGML,
        local_name="ext",
        attributes=[
            AttributeConstraint(
                namespace=None,
                local_name="cx",
                type_validator=get_type_validator(XsdBuiltinType.NON_NEGATIVE_INTEGER),
                required=True,
            ),
            AttributeConstraint(
                namespace=None,
                local_name="cy",
                type_validator=get_type_validator(XsdBuiltinType.NON_NEGATIVE_INTEGER),
                required=True,
            ),
        ],
    )
    registry.register(ext)

    return registry


# Global registry instance
PPTX_CONSTRAINTS = _create_pptx_constraints()


def get_element_constraint(namespace: str, local_name: str) -> ElementConstraint | None:
    """Get the constraint for an element.

    Args:
        namespace: The element namespace.
        local_name: The element local name.

    Returns:
        The ElementConstraint, or None if not found.
    """
    return PPTX_CONSTRAINTS.get_by_name(namespace, local_name)


def get_constraint_for_tag(tag: str) -> ElementConstraint | None:
    """Get constraint for a Clark notation tag.

    Args:
        tag: The element tag in Clark notation (e.g., "{namespace}localname").

    Returns:
        The ElementConstraint, or None if not found.
    """
    return PPTX_CONSTRAINTS.get(tag)
