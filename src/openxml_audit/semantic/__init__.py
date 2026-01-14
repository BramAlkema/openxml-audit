"""Semantic validation for Open XML documents."""

from openxml_audit.semantic.attributes import (
    AttributeMinMaxConstraint,
    AttributeMutualExclusive,
    AttributeRequiredConditionToValue,
    AttributeValueLengthConstraint,
    AttributeValueLessEqualToAnother,
    AttributeValuePatternConstraint,
    AttributeValueRangeConstraint,
    AttributeValueSetConstraint,
    SemanticConstraint,
)
from openxml_audit.semantic.references import (
    IdTracker,
    IndexReferenceConstraint,
    ReferenceExistConstraint,
    UniqueAttributeValueConstraint,
    validate_unique_ids,
)
from openxml_audit.semantic.relationships import (
    RelationshipExistConstraint,
    RelationshipTargetExistsConstraint,
    RelationshipTypeConstraint,
    validate_part_relationships,
)
from openxml_audit.semantic.validator import (
    SemanticValidator,
    create_pptx_semantic_validator,
)

__all__ = [
    # Base
    "SemanticConstraint",
    # Attributes
    "AttributeMinMaxConstraint",
    "AttributeMutualExclusive",
    "AttributeRequiredConditionToValue",
    "AttributeValueLengthConstraint",
    "AttributeValueLessEqualToAnother",
    "AttributeValuePatternConstraint",
    "AttributeValueRangeConstraint",
    "AttributeValueSetConstraint",
    # Relationships
    "RelationshipExistConstraint",
    "RelationshipTargetExistsConstraint",
    "RelationshipTypeConstraint",
    "validate_part_relationships",
    # References
    "IdTracker",
    "IndexReferenceConstraint",
    "ReferenceExistConstraint",
    "UniqueAttributeValueConstraint",
    "validate_unique_ids",
    # Validator
    "SemanticValidator",
    "create_pptx_semantic_validator",
]
