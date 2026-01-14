"""Semantic constraint classes for Open XML validation."""

from openxml_audit.semantic.constraints.compound import (
    AndConstraint,
    ConditionalConstraint,
    OrConstraint,
)
from openxml_audit.semantic.constraints.cross_part import CrossPartCountConstraint
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
    "AttributeComparisonConstraint",
    "AttributeEqualsConstraint",
    "AttributeNotEqualConstraint",
    "AttributesPresentConstraint",
]
