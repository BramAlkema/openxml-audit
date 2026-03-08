"""Semantic constraint classes for Open XML validation."""

from openxml_audit.semantic.constraints.compound import (
    AndConstraint,
    ConditionalConstraint,
    OrConstraint,
)
from openxml_audit.semantic.constraints.cross_part import (
    CrossPartCountConstraint,
    CrossPartReferenceConstraint,
)
from openxml_audit.semantic.constraints.equality import (
    AttributeComparisonConstraint,
    AttributeEqualsConstraint,
    AttributeNotEqualConstraint,
    AttributesPresentConstraint,
)

__all__ = [
    "AndConstraint",
    "ConditionalConstraint",
    "OrConstraint",
    "CrossPartCountConstraint",
    "CrossPartReferenceConstraint",
    "AttributeComparisonConstraint",
    "AttributeEqualsConstraint",
    "AttributeNotEqualConstraint",
    "AttributesPresentConstraint",
]
