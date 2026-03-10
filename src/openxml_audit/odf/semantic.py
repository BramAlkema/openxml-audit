"""ODF semantic-core validator — delegates to constraint registry."""

from __future__ import annotations

from lxml import etree

from openxml_audit.errors import ValidationError
from openxml_audit.odf.constraints import (
    EvaluationContext,
    build_default_registry,
)
from openxml_audit.odf.constraints.base import (
    ConstraintRegistry,
    OdfSemanticRule,
    detect_odf_version,
)
from openxml_audit.odf.package import OdfPackage


def get_odf_semantic_rules() -> tuple[OdfSemanticRule, ...]:
    """Return semantic-core rule metadata with stable identifiers."""
    return build_default_registry().rules()


class OdfSemanticValidator:
    """Semantic-core checks for ODF packages.

    Delegates all validation to the constraint registry.
    """

    def __init__(self, registry: ConstraintRegistry | None = None) -> None:
        self._registry = registry or build_default_registry()

    def validate(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        odf_version = detect_odf_version(parsed_parts)
        ctx = EvaluationContext(
            package=package,
            parsed_parts=parsed_parts,
            odf_version=odf_version,
        )
        return self._registry.evaluate_all(ctx)
