"""Particle validators for XML schema validation.

Particles define the structure of child elements:
- Sequence: elements must appear in order
- Choice: one of several options
- All: all elements required, any order
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext


class ParticleType(Enum):
    """Types of particles in XML Schema."""

    ELEMENT = "element"
    SEQUENCE = "sequence"
    CHOICE = "choice"
    ALL = "all"
    ANY = "any"
    GROUP = "group"


@dataclass
class ParticleConstraint:
    """Base constraint for particles."""

    particle_type: ParticleType
    min_occurs: int = 1
    max_occurs: int = 1  # -1 means unbounded
    namespace: str | None = None
    local_name: str | None = None

    @property
    def is_optional(self) -> bool:
        return self.min_occurs == 0

    @property
    def is_unbounded(self) -> bool:
        return self.max_occurs == -1


@dataclass
class ElementParticle(ParticleConstraint):
    """Particle for a specific element."""

    def __init__(
        self,
        namespace: str,
        local_name: str,
        min_occurs: int = 1,
        max_occurs: int = 1,
    ):
        super().__init__(
            particle_type=ParticleType.ELEMENT,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
            namespace=namespace,
            local_name=local_name,
        )

    @property
    def qualified_name(self) -> str:
        """Get the Clark notation qualified name."""
        if self.namespace:
            return f"{{{self.namespace}}}{self.local_name}"
        return self.local_name or ""


@dataclass
class CompositeParticle(ParticleConstraint):
    """Particle containing child particles (sequence, choice, all)."""

    children: list[ParticleConstraint] = field(default_factory=list)

    def add_child(self, child: ParticleConstraint) -> None:
        self.children.append(child)


@dataclass
class SequenceParticle(CompositeParticle):
    """Sequence particle - children must appear in order."""

    def __init__(
        self,
        children: list[ParticleConstraint] | None = None,
        min_occurs: int = 1,
        max_occurs: int = 1,
    ):
        super().__init__(
            particle_type=ParticleType.SEQUENCE,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
        )
        if children:
            self.children = children


@dataclass
class ChoiceParticle(CompositeParticle):
    """Choice particle - one of the children must appear."""

    def __init__(
        self,
        children: list[ParticleConstraint] | None = None,
        min_occurs: int = 1,
        max_occurs: int = 1,
    ):
        super().__init__(
            particle_type=ParticleType.CHOICE,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
        )
        if children:
            self.children = children


@dataclass
class AllParticle(CompositeParticle):
    """All particle - all children required but any order."""

    def __init__(
        self,
        children: list[ParticleConstraint] | None = None,
        min_occurs: int = 1,
        max_occurs: int = 1,
    ):
        super().__init__(
            particle_type=ParticleType.ALL,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
        )
        if children:
            self.children = children


@dataclass
class AnyParticle(ParticleConstraint):
    """Any particle - allows any element from namespace."""

    namespace_constraint: str = "##any"  # ##any, ##other, ##local, ##targetNamespace, or URI

    def __init__(
        self,
        namespace_constraint: str = "##any",
        min_occurs: int = 0,
        max_occurs: int = -1,
    ):
        super().__init__(
            particle_type=ParticleType.ANY,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
        )
        self.namespace_constraint = namespace_constraint


class ParticleValidator(ABC):
    """Base class for particle validators."""

    @abstractmethod
    def validate(
        self,
        constraint: ParticleConstraint,
        children: list[etree._Element],
        context: ValidationContext,
    ) -> bool:
        """Validate children against the particle constraint.

        Args:
            constraint: The particle constraint to validate against.
            children: The child elements to validate.
            context: The validation context.

        Returns:
            True if validation passed, False otherwise.
        """
        pass


class SequenceParticleValidator(ParticleValidator):
    """Validates sequence particles."""

    def validate(
        self,
        constraint: ParticleConstraint,
        children: list[etree._Element],
        context: ValidationContext,
    ) -> bool:
        if not isinstance(constraint, SequenceParticle):
            return False

        child_index = 0
        valid = True

        for particle in constraint.children:
            count = 0

            # Count matching elements
            while child_index < len(children):
                child = children[child_index]

                if self._matches(particle, child):
                    count += 1
                    child_index += 1

                    if particle.max_occurs != -1 and count >= particle.max_occurs:
                        break
                else:
                    break

            # Check occurrence constraints
            if count < particle.min_occurs:
                if isinstance(particle, ElementParticle):
                    context.add_schema_error(
                        f"Required element '{particle.local_name}' is missing "
                        f"(minOccurs={particle.min_occurs}, found={count})",
                        node=particle.local_name,
                    )
                valid = False

        # Check for unexpected elements
        if child_index < len(children):
            unexpected = children[child_index]
            tag = unexpected.tag
            if tag.startswith("{"):
                tag = tag.split("}")[-1]
            context.add_schema_error(
                f"Unexpected element '{tag}' found",
                node=tag,
            )
            valid = False

        return valid

    def _matches(self, particle: ParticleConstraint, element: etree._Element) -> bool:
        """Check if an element matches a particle."""
        if isinstance(particle, ElementParticle):
            return element.tag == particle.qualified_name
        elif isinstance(particle, AnyParticle):
            return self._matches_any(particle, element)
        elif isinstance(particle, CompositeParticle):
            # For composite particles, check if any child matches
            for child in particle.children:
                if self._matches(child, element):
                    return True
        return False

    def _matches_any(self, particle: AnyParticle, element: etree._Element) -> bool:
        """Check if element matches an any particle."""
        ns_constraint = particle.namespace_constraint

        if ns_constraint == "##any":
            return True
        elif ns_constraint == "##local":
            return not element.tag.startswith("{")
        elif ns_constraint == "##other":
            # Would need target namespace context
            return True
        else:
            # Specific namespace URI
            return element.tag.startswith(f"{{{ns_constraint}}}")


class ChoiceParticleValidator(ParticleValidator):
    """Validates choice particles."""

    def validate(
        self,
        constraint: ParticleConstraint,
        children: list[etree._Element],
        context: ValidationContext,
    ) -> bool:
        if not isinstance(constraint, ChoiceParticle):
            return False

        if not children:
            if constraint.min_occurs > 0:
                context.add_schema_error(
                    "Required choice element is missing",
                )
                return False
            return True

        # Check that first child matches one of the choices
        child = children[0]
        for particle in constraint.children:
            if self._matches(particle, child):
                return True

        # No match found
        tag = child.tag
        if tag.startswith("{"):
            tag = tag.split("}")[-1]

        expected = []
        for p in constraint.children:
            if isinstance(p, ElementParticle):
                expected.append(p.local_name or "")

        context.add_schema_error(
            f"Element '{tag}' is not a valid choice. "
            f"Expected one of: {', '.join(expected)}",
            node=tag,
        )
        return False

    def _matches(self, particle: ParticleConstraint, element: etree._Element) -> bool:
        """Check if an element matches a particle."""
        if isinstance(particle, ElementParticle):
            return element.tag == particle.qualified_name
        elif isinstance(particle, AnyParticle):
            return self._matches_any(particle, element)
        elif isinstance(particle, CompositeParticle):
            for child in particle.children:
                if self._matches(child, element):
                    return True
        return False

    def _matches_any(self, particle: AnyParticle, element: etree._Element) -> bool:
        """Check if element matches an any particle."""
        ns_constraint = particle.namespace_constraint

        if ns_constraint == "##any":
            return True
        elif ns_constraint == "##local":
            return not element.tag.startswith("{")
        elif ns_constraint == "##other":
            return True
        else:
            return element.tag.startswith(f"{{{ns_constraint}}}")


class AllParticleValidator(ParticleValidator):
    """Validates all particles (all elements required, any order)."""

    def validate(
        self,
        constraint: ParticleConstraint,
        children: list[etree._Element],
        context: ValidationContext,
    ) -> bool:
        if not isinstance(constraint, AllParticle):
            return False

        valid = True
        found: set[str] = set()

        # Track which elements we found
        for child in children:
            for particle in constraint.children:
                if isinstance(particle, ElementParticle):
                    if child.tag == particle.qualified_name:
                        if particle.qualified_name in found and particle.max_occurs == 1:
                            context.add_schema_error(
                                f"Duplicate element '{particle.local_name}' not allowed",
                                node=particle.local_name,
                            )
                            valid = False
                        found.add(particle.qualified_name)
                        break

        # Check all required elements present
        for particle in constraint.children:
            if isinstance(particle, ElementParticle):
                if particle.min_occurs > 0 and particle.qualified_name not in found:
                    context.add_schema_error(
                        f"Required element '{particle.local_name}' is missing",
                        node=particle.local_name,
                    )
                    valid = False

        return valid


def get_validator(particle_type: ParticleType) -> ParticleValidator | None:
    """Get the appropriate validator for a particle type."""
    validators: dict[ParticleType, ParticleValidator] = {
        ParticleType.SEQUENCE: SequenceParticleValidator(),
        ParticleType.CHOICE: ChoiceParticleValidator(),
        ParticleType.ALL: AllParticleValidator(),
    }
    return validators.get(particle_type)
