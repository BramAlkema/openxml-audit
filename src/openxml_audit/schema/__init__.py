"""Schema validation for Open XML documents."""

from openxml_audit.schema.constraints import (
    AttributeConstraint,
    ElementConstraint,
    ElementConstraintRegistry,
    PPTX_CONSTRAINTS,
    get_constraint_for_tag,
    get_element_constraint,
)
from openxml_audit.schema.particle import (
    AllParticle,
    AllParticleValidator,
    AnyParticle,
    ChoiceParticle,
    ChoiceParticleValidator,
    CompositeParticle,
    ElementParticle,
    ParticleConstraint,
    ParticleType,
    ParticleValidator,
    SequenceParticle,
    SequenceParticleValidator,
    get_validator,
)
from openxml_audit.schema.types import (
    BooleanTypeValidator,
    DecimalTypeValidator,
    IntegerTypeValidator,
    StringTypeValidator,
    TypeValidationResult,
    XsdBuiltinType,
    XsdTypeValidator,
    get_type_validator,
)
from openxml_audit.schema.validator import SchemaValidator

__all__ = [
    # Constraints
    "AttributeConstraint",
    "ElementConstraint",
    "ElementConstraintRegistry",
    "PPTX_CONSTRAINTS",
    "get_constraint_for_tag",
    "get_element_constraint",
    # Particles
    "AllParticle",
    "AllParticleValidator",
    "AnyParticle",
    "ChoiceParticle",
    "ChoiceParticleValidator",
    "CompositeParticle",
    "ElementParticle",
    "ParticleConstraint",
    "ParticleType",
    "ParticleValidator",
    "SequenceParticle",
    "SequenceParticleValidator",
    "get_validator",
    # Types
    "BooleanTypeValidator",
    "DecimalTypeValidator",
    "IntegerTypeValidator",
    "StringTypeValidator",
    "TypeValidationResult",
    "XsdBuiltinType",
    "XsdTypeValidator",
    "get_type_validator",
    # Validator
    "SchemaValidator",
]
