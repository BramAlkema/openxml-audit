"""Base constraint class for ODF semantic validation."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field

from lxml import etree

from openxml_audit.errors import ValidationError, ValidationErrorType, ValidationSeverity
from openxml_audit.odf._helpers import OFFICE_NS, normalize_part_uri
from openxml_audit.odf.package import OdfPackage


def parse_version_tuple(version: str | None) -> tuple[int, ...] | None:
    """Parse a version string like '1.3' or '1.3+csd01' into a tuple (1, 3)."""
    if version is None:
        return None
    stripped = version.strip()
    if not stripped:
        return None
    # Strip everything after first non-digit/dot character
    numeric = ""
    for ch in stripped:
        if ch.isdigit() or ch == ".":
            numeric += ch
        else:
            break
    if not numeric:
        return None
    parts: list[int] = []
    for part in numeric.rstrip(".").split("."):
        if part:
            parts.append(int(part))
    return tuple(parts) if parts else None


def detect_odf_version(parsed_parts: dict[str, etree._Element]) -> str | None:
    """Detect ODF version from office:version attribute on content.xml root."""
    content = parsed_parts.get("content.xml")
    if content is not None:
        version = content.get(f"{{{OFFICE_NS}}}version")
        if version:
            return version.strip()
    for part in ("styles.xml", "meta.xml", "settings.xml"):
        root = parsed_parts.get(part)
        if root is not None:
            version = root.get(f"{{{OFFICE_NS}}}version")
            if version:
                return version.strip()
    return None


@dataclass(frozen=True)
class OdfSemanticRule:
    """Stable semantic-core rule metadata."""

    id: str
    family: str
    description: str
    min_version: str | None = None
    max_version: str | None = None


@dataclass
class EvaluationContext:
    """Context passed to constraint evaluation."""

    package: OdfPackage
    parsed_parts: dict[str, etree._Element]
    odf_version: str | None = None
    _cache: dict[str, object] = field(default_factory=dict, repr=False)

    def cached(self, key: str, factory: object) -> object:
        """Return a cached value, computing it via *factory* on first access."""
        if key not in self._cache:
            self._cache[key] = factory()  # type: ignore[operator]
        return self._cache[key]


class OdfConstraint(ABC):
    """Base class for all ODF semantic constraints.

    Each constraint encapsulates a single validation rule with a stable ID.
    Subclasses may set ``min_version`` / ``max_version`` on their rule to
    restrict evaluation to specific ODF versions.
    """

    @property
    @abstractmethod
    def rule(self) -> OdfSemanticRule:
        """Return the rule metadata for this constraint."""

    @property
    def id(self) -> str:
        return self.rule.id

    @property
    def family(self) -> str:
        return self.rule.family

    @property
    def min_version(self) -> str | None:
        return self.rule.min_version

    @property
    def max_version(self) -> str | None:
        return self.rule.max_version

    def applies_to_version(self, odf_version: str | None) -> bool:
        """Check whether this constraint applies to the given ODF version.

        If the document version is unknown, constraints always apply.
        If the constraint has no version bounds, it always applies.
        """
        if odf_version is None:
            return True
        doc_ver = parse_version_tuple(odf_version)
        if doc_ver is None:
            return True

        min_ver = parse_version_tuple(self.min_version)
        if min_ver is not None and doc_ver < min_ver:
            return False

        max_ver = parse_version_tuple(self.max_version)
        return not (max_ver is not None and doc_ver > max_ver)

    @abstractmethod
    def evaluate(self, ctx: EvaluationContext) -> list[ValidationError]:
        """Evaluate this constraint and return any errors."""

    @staticmethod
    def _error(
        *,
        rule_id: str,
        error_type: ValidationErrorType,
        description: str,
        part_uri: str,
        severity: ValidationSeverity = ValidationSeverity.ERROR,
    ) -> ValidationError:
        return ValidationError(
            error_type=error_type,
            description=description,
            part_uri=part_uri,
            severity=severity,
            id=rule_id,
        )

    @staticmethod
    def _normalize_part_uri(part_path: str) -> str:
        return normalize_part_uri(part_path)


@dataclass
class ConstraintRegistry:
    """Registry of ODF constraints."""

    _constraints: list[OdfConstraint] = field(default_factory=list)

    def register(self, constraint: OdfConstraint) -> None:
        self._constraints.append(constraint)

    def rules(self) -> tuple[OdfSemanticRule, ...]:
        return tuple(c.rule for c in self._constraints)

    def evaluate_all(self, ctx: EvaluationContext) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for constraint in self._constraints:
            if not constraint.applies_to_version(ctx.odf_version):
                continue
            errors.extend(constraint.evaluate(ctx))
        return errors
