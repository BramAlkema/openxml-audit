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

    introduced_version: str | None = None

    def __init__(
        self,
        namespace: str,
        local_name: str,
        min_occurs: int = 1,
        max_occurs: int = 1,
        introduced_version: str | None = None,
    ):
        super().__init__(
            particle_type=ParticleType.ELEMENT,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
            namespace=namespace,
            local_name=local_name,
        )
        self.introduced_version = introduced_version

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
            consumed = self._consume(particle, children, child_index)

            if consumed == 0 and particle.min_occurs > 0:
                if isinstance(particle, ElementParticle):
                    context.add_schema_error(
                        f"Required element '{particle.local_name}' is missing "
                        f"(minOccurs={particle.min_occurs}, found=0)",
                        node=particle.local_name,
                    )
                else:
                    context.add_schema_error("Required sequence content is missing")
                valid = False
                continue

            child_index += consumed

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

    def _consume(
        self,
        particle: ParticleConstraint,
        children: list[etree._Element],
        start: int,
    ) -> int:
        if isinstance(particle, ElementParticle):
            return self._consume_element(particle, children, start)
        if isinstance(particle, AnyParticle):
            return self._consume_any(particle, children, start)
        if isinstance(particle, SequenceParticle):
            return self._consume_sequence(particle, children, start)
        if isinstance(particle, ChoiceParticle):
            return self._consume_choice(particle, children, start)
        return 0

    def _consume_element(
        self,
        particle: ElementParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        count = 0
        idx = start
        while idx < len(children):
            if children[idx].tag != particle.qualified_name:
                break
            count += 1
            idx += 1
            if particle.max_occurs != -1 and count >= particle.max_occurs:
                break
        if count < particle.min_occurs:
            return 0
        return idx - start

    def _consume_any(
        self,
        particle: AnyParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        count = 0
        idx = start
        while idx < len(children):
            if not self._matches_any(particle, children[idx]):
                break
            count += 1
            idx += 1
            if particle.max_occurs != -1 and count >= particle.max_occurs:
                break
        if count < particle.min_occurs:
            return 0
        return idx - start

    def _consume_sequence(
        self,
        particle: SequenceParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        total_consumed = 0
        occurrences = 0
        idx = start

        while particle.max_occurs == -1 or occurrences < particle.max_occurs:
            before = idx
            matched_once = True
            for child_particle in particle.children:
                consumed = self._consume(child_particle, children, idx)
                if consumed == 0 and child_particle.min_occurs > 0:
                    matched_once = False
                    break
                idx += consumed

            if not matched_once:
                idx = before
                break

            if idx == before:
                break

            occurrences += 1
            total_consumed += idx - before

        if occurrences < particle.min_occurs:
            return 0
        return total_consumed

    def _consume_choice(
        self,
        particle: ChoiceParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        total_consumed = 0
        occurrences = 0
        idx = start

        while particle.max_occurs == -1 or occurrences < particle.max_occurs:
            best = 0
            for option in particle.children:
                consumed = self._consume(option, children, idx)
                if consumed > best:
                    best = consumed
            if best == 0:
                break
            idx += best
            total_consumed += best
            occurrences += 1

        if occurrences < particle.min_occurs:
            return 0
        return total_consumed


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
            if constraint.min_occurs > 0 and not any(
                self._can_match_empty(particle) for particle in constraint.children
            ):
                context.add_schema_error(
                    "Required choice element is missing",
                )
                return False
            return True

        valid = True
        choice_count = 0
        child_index = 0
        expected = self._expected_names(constraint)

        while child_index < len(children):
            consumed = 0
            for particle in constraint.children:
                candidate = self._consume(particle, children, child_index)
                if candidate > consumed:
                    consumed = candidate

            if consumed > 0:
                choice_count += 1
                child_index += consumed
                continue

            child = children[child_index]
            tag = child.tag
            if tag.startswith("{"):
                tag = tag.split("}")[-1]

            context.add_schema_error(
                f"Element '{tag}' is not a valid choice. Expected one of: {', '.join(expected)}",
                node=tag,
            )
            valid = False
            child_index += 1

        if choice_count < constraint.min_occurs:
            context.add_schema_error(
                f"Choice requires at least {constraint.min_occurs} occurrence(s), "
                f"found {choice_count}",
            )
            valid = False

        if constraint.max_occurs != -1 and choice_count > constraint.max_occurs:
            context.add_schema_error(
                f"Choice allows at most {constraint.max_occurs} occurrence(s), "
                f"found {choice_count}",
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

    def _expected_names(self, constraint: ChoiceParticle) -> list[str]:
        expected = []
        for particle in constraint.children:
            if isinstance(particle, ElementParticle):
                expected.append(particle.local_name or "")
        return expected

    def _can_match_empty(self, particle: ParticleConstraint) -> bool:
        if particle.min_occurs == 0:
            return True
        if isinstance(particle, (ElementParticle, AnyParticle)):
            return False
        if isinstance(particle, ChoiceParticle):
            return any(self._can_match_empty(child) for child in particle.children)
        if isinstance(particle, CompositeParticle):
            return all(self._can_match_empty(child) for child in particle.children)
        return False

    def _consume(
        self,
        particle: ParticleConstraint,
        children: list[etree._Element],
        start: int,
    ) -> int:
        if isinstance(particle, ElementParticle):
            return self._consume_element(particle, children, start)
        if isinstance(particle, AnyParticle):
            return self._consume_any(particle, children, start)
        if isinstance(particle, SequenceParticle):
            return self._consume_sequence(particle, children, start)
        if isinstance(particle, ChoiceParticle):
            return self._consume_choice(particle, children, start)
        return 0

    def _consume_element(
        self,
        particle: ElementParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        count = 0
        idx = start
        while idx < len(children):
            if children[idx].tag != particle.qualified_name:
                break
            count += 1
            idx += 1
            if particle.max_occurs != -1 and count >= particle.max_occurs:
                break
        if count < particle.min_occurs:
            return 0
        return idx - start

    def _consume_any(
        self,
        particle: AnyParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        count = 0
        idx = start
        while idx < len(children):
            if not self._matches_any(particle, children[idx]):
                break
            count += 1
            idx += 1
            if particle.max_occurs != -1 and count >= particle.max_occurs:
                break
        if count < particle.min_occurs:
            return 0
        return idx - start

    def _consume_sequence(
        self,
        particle: SequenceParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        total_consumed = 0
        occurrences = 0
        idx = start

        while particle.max_occurs == -1 or occurrences < particle.max_occurs:
            before = idx
            matched_once = True
            for child_particle in particle.children:
                consumed = self._consume(child_particle, children, idx)
                if consumed == 0 and child_particle.min_occurs > 0:
                    matched_once = False
                    break
                idx += consumed

            if not matched_once:
                idx = before
                break

            if idx == before:
                break

            occurrences += 1
            total_consumed += idx - before

        if occurrences < particle.min_occurs:
            return 0
        return total_consumed

    def _consume_choice(
        self,
        particle: ChoiceParticle,
        children: list[etree._Element],
        start: int,
    ) -> int:
        total_consumed = 0
        occurrences = 0
        idx = start

        while particle.max_occurs == -1 or occurrences < particle.max_occurs:
            best = 0
            for option in particle.children:
                consumed = self._consume(option, children, idx)
                if consumed > best:
                    best = consumed
            if best == 0:
                break
            idx += best
            total_consumed += best
            occurrences += 1

        if occurrences < particle.min_occurs:
            return 0
        return total_consumed


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
        counts: list[int] = [0] * len(constraint.children)

        for child in children:
            matched = False
            for idx, particle in enumerate(constraint.children):
                if not self._matches(particle, child):
                    continue
                matched = True
                counts[idx] += 1
                if particle.max_occurs != -1 and counts[idx] > particle.max_occurs:
                    context.add_schema_error(
                        f"Element '{self._particle_name(particle)}' exceeds "
                        f"maxOccurs={particle.max_occurs}",
                        node=self._particle_name(particle),
                    )
                    valid = False
                break

            if not matched:
                tag = child.tag
                if tag.startswith("{"):
                    tag = tag.split("}")[-1]
                context.add_schema_error(
                    f"Unexpected element '{tag}' found",
                    node=tag,
                )
                valid = False

        for idx, particle in enumerate(constraint.children):
            if counts[idx] < particle.min_occurs:
                context.add_schema_error(
                    f"Required element '{self._particle_name(particle)}' is missing "
                    f"(minOccurs={particle.min_occurs}, found={counts[idx]})",
                    node=self._particle_name(particle),
                )
                valid = False

        return valid

    def _matches(self, particle: ParticleConstraint, element: etree._Element) -> bool:
        if isinstance(particle, ElementParticle):
            return element.tag == particle.qualified_name
        if isinstance(particle, AnyParticle):
            return self._matches_any(particle, element)
        if isinstance(particle, CompositeParticle):
            return any(self._matches(child, element) for child in particle.children)
        return False

    def _matches_any(self, particle: AnyParticle, element: etree._Element) -> bool:
        ns_constraint = particle.namespace_constraint

        if ns_constraint == "##any":
            return True
        if ns_constraint == "##local":
            return not element.tag.startswith("{")
        if ns_constraint == "##other":
            return True
        return element.tag.startswith(f"{{{ns_constraint}}}")

    def _particle_name(self, particle: ParticleConstraint) -> str:
        if isinstance(particle, ElementParticle):
            return particle.local_name or "element"
        if isinstance(particle, AnyParticle):
            return "any"
        return particle.particle_type.value


_VALIDATORS: dict[ParticleType, ParticleValidator] = {
    ParticleType.SEQUENCE: SequenceParticleValidator(),
    ParticleType.CHOICE: ChoiceParticleValidator(),
    ParticleType.ALL: AllParticleValidator(),
}


def get_validator(particle_type: ParticleType) -> ParticleValidator | None:
    """Get the appropriate validator for a particle type."""
    return _VALIDATORS.get(particle_type)
