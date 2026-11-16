"""Shared helpers for whole-dataset bridge invariant tests."""

from __future__ import annotations

import json
from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal
from functools import lru_cache

from lxml import etree

from openxml_audit.codegen.data_resources import get_openxml_data_dir
from openxml_audit.codegen.schema_loader import SdkAttribute, get_registry
from openxml_audit.context import ValidationContext
from openxml_audit.errors import FileFormat
from openxml_audit.schema.types import (
    AnyURITypeValidator,
    DecimalTypeValidator,
    IntegerTypeValidator,
    ListTypeValidator,
    NCNameTypeValidator,
    QNameTypeValidator,
    StringTypeValidator,
    UnionTypeValidator,
    VersionedTypeValidator,
    XsdTypeValidator,
)

_ENUM_MARKER = "EnumValue<"
_VERSIONED_PROBE_CANDIDATES = (
    "0",
    "1",
    "2",
    "10",
    "25",
    "100",
    "-1",
    "-2",
    "0.5",
    "1.5",
    "25%",
    "100%",
    "-25%",
    "indefinite",
    "normal",
    "warning",
    "error",
    "priority",
    "basedOn",
    "default",
    "0000",
    "0001",
    "ab",
    "abc",
    "foo",
    "bar",
    "__invalid__",
)


@dataclass(frozen=True)
class LiveEnumAttributeUse:
    """Live enum-bearing attribute from shipped schema data."""

    element_type_name: str
    attribute_qname: str
    sdk_type_name: str
    enum_type_names: tuple[str, ...]
    has_enum_validator: bool
    expects_union_wrapper: bool
    expects_versioned_wrapper: bool

    @property
    def owner(self) -> str:
        return f"{self.element_type_name} {self.attribute_qname}"


@dataclass(frozen=True)
class LiveNumberValidatorUse:
    """Live NumberValidator-bearing attribute branch from shipped schema data."""

    element_type_name: str
    attribute_qname: str
    sdk_type_name: str
    validator_index: int
    validator_type_name: str | None
    version: str | None
    union_id: int | None
    is_list: bool
    args: tuple[tuple[str, str], ...]

    @property
    def owner(self) -> str:
        return f"{self.element_type_name} {self.attribute_qname}"


@dataclass(frozen=True)
class LiveStringValidatorUse:
    """Live StringValidator-bearing attribute branch from shipped schema data."""

    element_type_name: str
    attribute_qname: str
    sdk_type_name: str
    validator_index: int
    validator_type_name: str | None
    version: str | None
    union_id: int | None
    args: tuple[tuple[str, str], ...]

    @property
    def owner(self) -> str:
        return f"{self.element_type_name} {self.attribute_qname}"


@dataclass(frozen=True)
class LiveVersionedValidatorUse:
    """Live attribute whose runtime bridge produces a VersionedTypeValidator."""

    element_type_name: str
    attribute_qname: str
    sdk_type_name: str
    branch_versions: tuple[str | None, ...]

    @property
    def owner(self) -> str:
        return f"{self.element_type_name} {self.attribute_qname}"


def extract_enum_type_names(sdk_type: str) -> tuple[str, ...]:
    """Extract all EnumValue<T> type names from an SDK type string."""
    enum_types: list[str] = []
    start = sdk_type.find(_ENUM_MARKER)
    while start != -1:
        inner = sdk_type[start + len(_ENUM_MARKER) :]
        depth = 1
        value: list[str] = []
        for char in inner:
            if char == "<":
                depth += 1
            elif char == ">":
                depth -= 1
                if depth == 0:
                    break
            value.append(char)
        enum_types.append("".join(value))
        start = sdk_type.find(_ENUM_MARKER, start + len(_ENUM_MARKER))
    return tuple(enum_types)


@lru_cache(maxsize=1)
def collect_live_enum_type_names() -> tuple[str, ...]:
    """Collect all live .NET EnumValue<T> uses from shipped schema JSON."""
    data_dir = get_openxml_data_dir()
    enum_types: set[str] = set()

    def walk_types(payload: object) -> None:
        if isinstance(payload, dict):
            sdk_type = payload.get("Type")
            if isinstance(sdk_type, str):
                for enum_type in extract_enum_type_names(sdk_type):
                    if enum_type.startswith("DocumentFormat.OpenXml."):
                        enum_types.add(enum_type)
            for child in payload.values():
                walk_types(child)
            return

        if isinstance(payload, list):
            for child in payload:
                walk_types(child)

    for schema_path in sorted((data_dir / "schemas").glob("*.json")):
        walk_types(json.loads(schema_path.read_text(encoding="utf-8")))

    return tuple(sorted(enum_types))


@lru_cache(maxsize=1)
def collect_live_enum_attribute_uses() -> tuple[LiveEnumAttributeUse, ...]:
    """Collect shipped attributes that directly use enum types or enum validators."""
    registry = get_registry()
    registry.load()
    uses: list[LiveEnumAttributeUse] = []

    for schema in registry._schemas.values():
        for elem in schema.types:
            for attr in elem.attributes:
                enum_type_names = tuple(
                    enum_type
                    for enum_type in extract_enum_type_names(attr.type_name)
                    if enum_type.startswith("DocumentFormat.OpenXml.")
                )
                filtered_validators = [
                    validator
                    for validator in attr.validators
                    if validator.get("Name") not in {"RequiredValidator", "OfficeVersionValidator"}
                ]
                has_enum_validator = any(
                    validator.get("Name") == "EnumValidator" for validator in filtered_validators
                )
                union_groups: dict[int, int] = {}
                for validator in filtered_validators:
                    union_id = validator.get("UnionId")
                    if union_id is None:
                        continue
                    union_groups[union_id] = union_groups.get(union_id, 0) + 1

                expects_union_wrapper = len(union_groups) > 1 or any(
                    count > 1 for count in union_groups.values()
                )
                expects_versioned_wrapper = any(
                    validator.get("Version") is not None for validator in filtered_validators
                )

                if not enum_type_names and not has_enum_validator:
                    continue

                uses.append(
                    LiveEnumAttributeUse(
                        element_type_name=elem.name,
                        attribute_qname=attr.qname,
                        sdk_type_name=attr.type_name,
                        enum_type_names=enum_type_names,
                        has_enum_validator=has_enum_validator,
                        expects_union_wrapper=expects_union_wrapper,
                        expects_versioned_wrapper=expects_versioned_wrapper,
                    )
                )

    return tuple(uses)


@lru_cache(maxsize=1)
def collect_live_enum_type_owners() -> dict[str, tuple[str, ...]]:
    owners: dict[str, set[str]] = defaultdict(set)
    for use in collect_live_enum_attribute_uses():
        for enum_type in use.enum_type_names:
            owners[enum_type].add(use.owner)
    return {enum_type: tuple(sorted(enum_owners)) for enum_type, enum_owners in owners.items()}


def get_attribute_for_use(
    use: (
        LiveEnumAttributeUse
        | LiveNumberValidatorUse
        | LiveStringValidatorUse
        | LiveVersionedValidatorUse
    ),
) -> SdkAttribute:
    registry = get_registry()
    elem = registry.get_type(use.element_type_name)
    if elem is None:
        raise AssertionError(
            f"Unknown SDK element type in invariant helper: {use.element_type_name}"
        )

    for attr in elem.attributes:
        if attr.qname == use.attribute_qname and attr.type_name == use.sdk_type_name:
            return attr

    raise AssertionError(
        "Unknown SDK attribute in invariant helper: "
        f"{use.element_type_name} {use.attribute_qname} ({use.sdk_type_name})"
    )


def get_validator_dict_for_use(
    use: LiveNumberValidatorUse | LiveStringValidatorUse,
) -> dict:
    attr = get_attribute_for_use(use)
    try:
        return attr.validators[use.validator_index]
    except IndexError as exc:
        raise AssertionError(
            "Unknown SDK validator in invariant helper: "
            f"{use.element_type_name} {use.attribute_qname} "
            f"validator_index={use.validator_index}"
        ) from exc


@lru_cache(maxsize=1)
def collect_live_number_validator_uses() -> tuple[LiveNumberValidatorUse, ...]:
    """Collect all shipped NumberValidator branches with owning attribute metadata."""
    registry = get_registry()
    registry.load()
    uses: list[LiveNumberValidatorUse] = []

    for schema in registry._schemas.values():
        for elem in schema.types:
            for attr in elem.attributes:
                for index, validator in enumerate(attr.validators):
                    if validator.get("Name") != "NumberValidator":
                        continue

                    args = tuple(
                        (arg["Name"], arg["Value"])
                        for arg in validator.get("Arguments", [])
                        if arg.get("Name") and arg.get("Value") is not None
                    )
                    uses.append(
                        LiveNumberValidatorUse(
                            element_type_name=elem.name,
                            attribute_qname=attr.qname,
                            sdk_type_name=attr.type_name,
                            validator_index=index,
                            validator_type_name=validator.get("Type"),
                            version=validator.get("Version"),
                            union_id=validator.get("UnionId"),
                            is_list=validator.get("IsList") in (True, "True"),
                            args=args,
                        )
                    )

    return tuple(uses)


@lru_cache(maxsize=1)
def partition_live_number_validator_uses() -> Counter[str]:
    """Partition shipped NumberValidator branches by runtime scalar kind and bounds."""
    partitions: Counter[str] = Counter()
    for use in collect_live_number_validator_uses():
        validator = build_number_validator_for_use(use)
        item_validator = unwrap_list_item_validator(validator)

        if isinstance(item_validator, IntegerTypeValidator):
            partitions["integer"] += 1
        elif isinstance(item_validator, DecimalTypeValidator):
            partitions["decimal"] += 1
        else:
            partitions[type(item_validator).__name__] += 1

        if use.is_list:
            partitions["list"] += 1
        if isinstance(item_validator, (IntegerTypeValidator, DecimalTypeValidator)):
            if item_validator.min_value is not None:
                partitions["min_bounded"] += 1
            if item_validator.max_value is not None:
                partitions["max_bounded"] += 1
            if item_validator.min_value is not None and Decimal(str(item_validator.min_value)) >= 0:
                partitions["nonnegative"] += 1

    return partitions


@lru_cache(maxsize=1)
def collect_live_string_validator_uses() -> tuple[LiveStringValidatorUse, ...]:
    """Collect all shipped StringValidator branches with owning attribute metadata."""
    registry = get_registry()
    registry.load()
    uses: list[LiveStringValidatorUse] = []

    for schema in registry._schemas.values():
        for elem in schema.types:
            for attr in elem.attributes:
                for index, validator in enumerate(attr.validators):
                    if validator.get("Name") != "StringValidator":
                        continue

                    args = tuple(
                        (arg["Name"], arg["Value"])
                        for arg in validator.get("Arguments", [])
                        if arg.get("Name") and arg.get("Value") is not None
                    )
                    uses.append(
                        LiveStringValidatorUse(
                            element_type_name=elem.name,
                            attribute_qname=attr.qname,
                            sdk_type_name=attr.type_name,
                            validator_index=index,
                            validator_type_name=validator.get("Type"),
                            version=validator.get("Version"),
                            union_id=validator.get("UnionId"),
                            args=args,
                        )
                    )

    return tuple(uses)


@lru_cache(maxsize=1)
def collect_live_versioned_validator_uses() -> tuple[LiveVersionedValidatorUse, ...]:
    """Collect all shipped attributes whose bridge yields a VersionedTypeValidator."""
    from openxml_audit.codegen.constraint_bridge import _build_type_validator

    registry = get_registry()
    registry.load()
    uses: list[LiveVersionedValidatorUse] = []

    for schema in registry._schemas.values():
        for elem in schema.types:
            for attr in elem.attributes:
                validator = _build_type_validator(attr)
                if not isinstance(validator, VersionedTypeValidator):
                    continue

                uses.append(
                    LiveVersionedValidatorUse(
                        element_type_name=elem.name,
                        attribute_qname=attr.qname,
                        sdk_type_name=attr.type_name,
                        branch_versions=tuple(version for version, _branch in validator.branches),
                    )
                )

    return tuple(uses)


def build_number_validator_for_use(use: LiveNumberValidatorUse) -> XsdTypeValidator:
    from openxml_audit.codegen.constraint_bridge import (
        _build_number_validator,
        _parse_validator_args,
    )

    validator_dict = get_validator_dict_for_use(use)
    return _build_number_validator(
        _parse_validator_args(validator_dict),
        get_attribute_for_use(use),
        validator_dict,
    )


def flatten_runtime_validators(validator: XsdTypeValidator | None) -> tuple[XsdTypeValidator, ...]:
    if validator is None:
        return ()
    if isinstance(validator, UnionTypeValidator):
        flattened: list[XsdTypeValidator] = []
        for member in validator.members:
            flattened.extend(flatten_runtime_validators(member))
        return tuple(flattened)
    if isinstance(validator, VersionedTypeValidator):
        flattened = []
        for _version, branch in validator.branches:
            flattened.extend(flatten_runtime_validators(branch))
        return tuple(flattened)
    return (validator,)


def unwrap_list_item_validator(validator: XsdTypeValidator) -> XsdTypeValidator:
    if isinstance(validator, ListTypeValidator):
        return validator.item_validator
    return validator


def find_fractional_decimal_probe(validator: DecimalTypeValidator) -> str | None:
    candidates = [
        Decimal("-10.5"),
        Decimal("-1.5"),
        Decimal("-0.5"),
        Decimal("0.1"),
        Decimal("0.5"),
        Decimal("0.9"),
        Decimal("1.1"),
        Decimal("1.5"),
        Decimal("2.5"),
        Decimal("10.5"),
    ]
    for candidate in candidates:
        if validator.validate(str(candidate)).is_valid:
            return str(candidate)

    if validator.min_value is not None and validator.max_value is not None:
        low = validator.min_value + (Decimal("0.1") if validator.min_inclusive else Decimal("0.2"))
        high = validator.max_value - (Decimal("0.1") if validator.max_inclusive else Decimal("0.2"))
        midpoint = (low + high) / 2
        if midpoint != int(midpoint) and validator.validate(str(midpoint)).is_valid:
            return str(midpoint.normalize())

    if validator.min_value is not None:
        candidate = validator.min_value + (
            Decimal("0.1") if validator.min_inclusive else Decimal("0.2")
        )
        if validator.validate(str(candidate)).is_valid:
            return str(candidate.normalize())

    if validator.max_value is not None:
        candidate = validator.max_value - (
            Decimal("0.1") if validator.max_inclusive else Decimal("0.2")
        )
        if validator.validate(str(candidate)).is_valid:
            return str(candidate.normalize())

    return None


def find_valid_numeric_probe(
    validator: IntegerTypeValidator | DecimalTypeValidator,
) -> str | None:
    if isinstance(validator, DecimalTypeValidator):
        probe = find_fractional_decimal_probe(validator)
        if probe is not None:
            return probe

        for candidate in ["0", "1", "-1", "2", "-2", "10"]:
            if validator.validate(candidate).is_valid:
                return candidate

        if validator.min_value is not None:
            candidate = Decimal(str(validator.min_value))
            if not validator.min_inclusive:
                candidate += Decimal("1")
            if validator.validate(str(candidate)).is_valid:
                return str(candidate.normalize())

        if validator.max_value is not None:
            candidate = Decimal(str(validator.max_value))
            if not validator.max_inclusive:
                candidate -= Decimal("1")
            if validator.validate(str(candidate)).is_valid:
                return str(candidate.normalize())

        return None

    for candidate in ["1", "0", "2", "-1", "10", "-10"]:
        if validator.validate(candidate).is_valid:
            return candidate

    if validator.min_value is not None:
        candidate = validator.min_value if validator.min_inclusive else validator.min_value + 1
        if validator.validate(str(candidate)).is_valid:
            return str(candidate)

    if validator.max_value is not None:
        candidate = validator.max_value if validator.max_inclusive else validator.max_value - 1
        if validator.validate(str(candidate)).is_valid:
            return str(candidate)

    return None


def get_lower_bound_rejection_probe(
    validator: IntegerTypeValidator | DecimalTypeValidator,
) -> str | None:
    if validator.min_value is None:
        return None
    if isinstance(validator, IntegerTypeValidator):
        return str(validator.min_value if not validator.min_inclusive else validator.min_value - 1)
    min_value = Decimal(str(validator.min_value))
    return str(min_value if not validator.min_inclusive else min_value - Decimal("0.1"))


def get_upper_bound_rejection_probe(
    validator: IntegerTypeValidator | DecimalTypeValidator,
) -> str | None:
    if validator.max_value is None:
        return None
    if isinstance(validator, IntegerTypeValidator):
        return str(validator.max_value if not validator.max_inclusive else validator.max_value + 1)
    max_value = Decimal(str(validator.max_value))
    return str(max_value if not validator.max_inclusive else max_value + Decimal("0.1"))


def build_string_validator_for_use(use: LiveStringValidatorUse) -> XsdTypeValidator | None:
    from openxml_audit.codegen.constraint_bridge import _build_type_validator

    return _build_type_validator(get_attribute_for_use(use))


def build_versioned_validator_for_use(use: LiveVersionedValidatorUse) -> VersionedTypeValidator:
    from openxml_audit.codegen.constraint_bridge import _build_type_validator

    validator = _build_type_validator(get_attribute_for_use(use))
    if not isinstance(validator, VersionedTypeValidator):
        raise AssertionError(f"{use.owner} no longer builds VersionedTypeValidator")
    return validator


def args_dict(
    use: LiveNumberValidatorUse | LiveStringValidatorUse,
) -> dict[str, str]:
    return dict(use.args)


def runtime_has_string_pattern(
    validator: XsdTypeValidator | None,
    pattern: str,
) -> bool:
    for branch in flatten_runtime_validators(validator):
        if (
            isinstance(branch, StringTypeValidator)
            and branch.pattern is not None
            and branch.pattern.pattern == pattern
        ):
            return True
    return False


def runtime_has_specialized_string_validator(
    validator: XsdTypeValidator | None,
    validator_type: (
        type[QNameTypeValidator] | type[NCNameTypeValidator] | type[AnyURITypeValidator]
    ),
) -> bool:
    return any(
        isinstance(branch, validator_type) for branch in flatten_runtime_validators(validator)
    )


def expected_selected_branch_for_format(
    validator: VersionedTypeValidator,
    file_format: FileFormat,
) -> tuple[str | None, XsdTypeValidator]:
    selected: tuple[str | None, XsdTypeValidator] | None = None
    for version, branch in validator.branches:
        if version is None:
            selected = (version, branch)
            continue
        introduced = FileFormat.from_version_string(version)
        if introduced is None or file_format.includes_ooxml(introduced):
            selected = (version, branch)
    if selected is not None:
        return selected
    return validator.branches[0]


def get_selected_branch_version_for_format(
    validator: VersionedTypeValidator,
    file_format: FileFormat,
) -> str | None:
    selected_branch = validator._select_validator(file_format)
    for version, branch in validator.branches:
        if branch is selected_branch:
            return version
    raise AssertionError(
        f"Selected branch for {file_format} is not present in VersionedTypeValidator"
    )


def versioned_probe_matrix(
    validator: VersionedTypeValidator,
    *,
    file_formats: tuple[FileFormat, ...] = (
        FileFormat.OFFICE_2007,
        FileFormat.OFFICE_2010,
        FileFormat.OFFICE_2013,
    ),
) -> dict[str, tuple[bool, ...]]:
    matrix: dict[str, tuple[bool, ...]] = {}
    for probe in _VERSIONED_PROBE_CANDIDATES:
        matrix[probe] = tuple(
            validator.validate(probe, ValidationContext(file_format=file_format)).is_valid
            for file_format in file_formats
        )
    return matrix


def find_transition_probe(
    validator: VersionedTypeValidator,
    older_format: FileFormat,
    newer_format: FileFormat,
) -> str | None:
    for probe, results in versioned_probe_matrix(
        validator,
        file_formats=(older_format, newer_format),
    ).items():
        if results == (False, True):
            return probe
    return None


def validator_has_enumeration_semantics(validator: XsdTypeValidator | None) -> bool:
    if validator is None:
        return False
    if isinstance(validator, StringTypeValidator):
        return validator.enumeration is not None
    if isinstance(validator, ListTypeValidator):
        return validator_has_enumeration_semantics(validator.item_validator)
    if isinstance(validator, UnionTypeValidator):
        return any(validator_has_enumeration_semantics(member) for member in validator.members)
    if isinstance(validator, VersionedTypeValidator):
        return any(
            validator_has_enumeration_semantics(branch_validator)
            for _version, branch_validator in validator.branches
        )
    return False


@lru_cache(maxsize=1)
def collect_ambiguous_element_candidates() -> dict[str, tuple[str, ...]]:
    """Collect element tags with multiple SDK type candidates."""
    registry = get_registry()
    registry.load()
    ambiguous: dict[str, tuple[str, ...]] = {}

    for tag, candidates in registry._elements_by_tag.items():
        if len(candidates) <= 1:
            continue
        ambiguous[tag] = tuple(candidate.class_name for candidate in candidates)

    return ambiguous


def get_selected_candidate_class_name(element: etree._Element) -> str | None:
    """Resolve the selected SDK candidate class name for a concrete element."""
    from openxml_audit.codegen.constraint_bridge import _get_sdk_element_type_for_element

    elem_type = _get_sdk_element_type_for_element(element.tag, element)
    return elem_type.class_name if elem_type is not None else None
