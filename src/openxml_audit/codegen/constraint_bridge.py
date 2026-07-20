"""Bridge between SDK schema data and runtime constraint validation.

Converts SDK types to our constraint classes on-demand.
"""

from __future__ import annotations

from decimal import Decimal
from functools import lru_cache
from typing import TYPE_CHECKING

from lxml import etree

from openxml_audit.codegen.schema_loader import (
    SdkAttribute,
    SdkElementType,
    SdkParticle,
    _extract_enum_type_name,
    _extract_list_item_type_name,
    get_enum_values,
    get_registry,
    get_xsd_type_name,
)
from openxml_audit.namespaces import (
    DRAWINGML_CHART,
    OFFICE_DOC_MATH,
    SPREADSHEETML,
    WORDPROCESSINGML,
)
from openxml_audit.schema.constraints import (
    AttributeConstraint,
    ElementConstraint,
)
from openxml_audit.schema.particle import (
    AllParticle,
    AnyParticle,
    ChoiceParticle,
    CompositeParticle,
    ElementParticle,
    ParticleConstraint,
    SequenceParticle,
)
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
    get_type_validator,
)

if TYPE_CHECKING:
    pass

# Cache for tags where get_element_type_candidates() returns exactly 1 candidate.
# The result is deterministic by tag alone (no dependency on element content).
_single_candidate_cache: dict[str, ElementConstraint] = {}

# Cache for tags with multiple candidates. The chosen candidate is fully
# determined by the structural signature the resolver actually reads
# (see _multi_candidate_key), and the SDK registry is static for the process,
# so the result is reusable across documents. Distinct signatures are few in
# practice (a few hundred even for large documents).
_MULTI_CACHE_MISS = object()
_multi_candidate_cache: dict[tuple, ElementConstraint | None] = {}

_CUSTOM_NUMBER_TYPE_VALIDATORS: dict[str, XsdTypeValidator] = {
    "a:ST_Angle": IntegerTypeValidator(),
    "a:ST_Coordinate": IntegerTypeValidator(),
    "a:ST_DrawingElementId": IntegerTypeValidator(min_value=0),
    "a:ST_PositiveFixedPercentage": IntegerTypeValidator(min_value=1),
    "msink:ST_Point": IntegerTypeValidator(),
    "w:ST_DecimalNumber": DecimalTypeValidator(),
    "w:ST_HpsMeasure_O12": IntegerTypeValidator(),
    "w:ST_NonNegativeDecimalNumber": DecimalTypeValidator(min_value=0),
    "w:ST_SignedDecimalNumberMax-1": DecimalTypeValidator(max_value=-1),
    "w:ST_SignedDecimalNumberMax-2": DecimalTypeValidator(max_value=-2),
    "w:ST_SignedHpsMeasure_O12": IntegerTypeValidator(),
    "w:ST_SignedTwipsMeasure_O12": IntegerTypeValidator(),
    "w:ST_TwipsMeasure_O12": IntegerTypeValidator(),
    "w:ST_UnsignedDecimalNumber": DecimalTypeValidator(min_value=0),
    "w:ST_UnsignedDecimalNumberMin1": DecimalTypeValidator(min_value=1),
}

_CHART_SERIES_CANDIDATE_BY_PARENT: dict[str, str] = {
    "areaChart": "AreaChartSeries",
    "area3DChart": "AreaChartSeries",
    "barChart": "BarChartSeries",
    "bar3DChart": "BarChartSeries",
    "bubbleChart": "BubbleChartSeries",
    "doughnutChart": "PieChartSeries",
    "lineChart": "LineChartSeries",
    "line3DChart": "LineChartSeries",
    "ofPieChart": "PieChartSeries",
    "pieChart": "PieChartSeries",
    "pie3DChart": "PieChartSeries",
    "radarChart": "RadarChartSeries",
    "scatterChart": "ScatterChartSeries",
    "stockChart": "LineChartSeries",
    "surfaceChart": "SurfaceChartSeries",
    "surface3DChart": "SurfaceChartSeries",
}

_CHART_EXTENSION_LIST_CANDIDATE_BY_PARENT: dict[str, str] = {
    "areaChart": "AreaChartExtensionList",
    "area3DChart": "Area3DChartExtensionList",
    "barChart": "BarChartExtensionList",
    "bar3DChart": "Bar3DChartExtensionList",
    "bubbleChart": "BubbleChartExtensionList",
    "chartSpace": "ChartSpaceExtensionList",
    "dLbls": "DLblsExtensionList",
    "lineChart": "LineChartExtensionList",
    "line3DChart": "Line3DChartExtensionList",
    "pieChart": "PieChartExtensionList",
    "pie3DChart": "Pie3DChartExtensionList",
    "radarChart": "RadarChartExtensionList",
    "scatterChart": "ScatterChartExtensionList",
    "stockChart": "StockChartExtensionList",
    "surfaceChart": "SurfaceChartExtensionList",
    "surface3DChart": "Surface3DChartExtensionList",
}

_CHART_EXTENSION_CANDIDATE_BY_PARENT: dict[str, str] = {
    "areaChart": "AreaChartExtension",
    "area3DChart": "Area3DChartExtension",
    "barChart": "BarChartExtension",
    "bar3DChart": "Bar3DChartExtension",
    "bubbleChart": "BubbleChartExtension",
    "chartSpace": "ChartSpaceExtension",
    "dLbls": "DLblsExtension",
    "lineChart": "LineChartExtension",
    "line3DChart": "Line3DChartExtension",
    "pieChart": "PieChartExtension",
    "pie3DChart": "Pie3DChartExtension",
    "radarChart": "RadarChartExtension",
    "scatterChart": "ScatterChartExtension",
    "stockChart": "StockChartExtension",
    "surfaceChart": "SurfaceChartExtension",
    "surface3DChart": "Surface3DChartExtension",
}

_SPREADSHEET_EXTENSION_LIST_CANDIDATE_BY_PARENT: dict[str, str] = {
    "queryTable": "QueryTableExtensionList",
    "worksheet": "WorksheetExtensionList",
}

_SPREADSHEET_EXTENSION_CANDIDATE_BY_PARENT: dict[str, str] = {
    "queryTable": "QueryTableExtension",
    "worksheet": "WorksheetExtension",
}

_WORD_RUN_PROPERTIES_CANDIDATE_BY_PARENT: dict[str, str] = {
    "pPr": "ParagraphMarkRunProperties",
    "rPrChange": "PreviousParagraphMarkRunProperties",
    "style": "StyleRunProperties",
}

_WORD_PARAGRAPH_PROPERTIES_CANDIDATE_BY_PARENT: dict[str, str] = {
    "lvl": "PreviousParagraphProperties",
    "p": "ParagraphProperties",
    "pPrChange": "ParagraphPropertiesExtended",
    "style": "StyleParagraphProperties",
    "tblStylePr": "StyleParagraphProperties",
}

_BAR_CHART_LOCALS = {"barChart", "bar3DChart"}

# w:start is three-way overloaded and w:end two-way: a border, a table cell
# margin, or (start only) a numbering start value. They share an element name
# and differ only in declaring complex type, so the parent decides.
#
# Scoring cannot separate them: `<w:start w:val="1"/>` matches both CT_Border
# and CT_NonNegativeDecimalNumber on every component of _score_candidate, and
# the final tiebreak prefers whichever declares MORE attributes — CT_Border,
# with nine. That picked the border type for numbering values, validating
# `w:val` against the page-border-art enumeration.
_WORD_START_END_TYPE_BY_PARENT = {
    "lvl": "CT_NonNegativeDecimalNumber",
    "pBdr": "CT_Border",
    "pgBorders": "CT_Border",
    "tblBorders": "CT_Border",
    "tcBorders": "CT_Border",
    "tblCellMar": "CT_TblWidth",
    "tcMar": "CT_TblWidth",
}


def _convert_attribute(
    attr: SdkAttribute,
    namespace_map: dict[str, str],
) -> AttributeConstraint:
    """Convert an SDK attribute to an AttributeConstraint."""
    # Determine namespace from prefix
    ns = None
    if attr.prefix:
        ns = namespace_map.get(attr.prefix)

    type_validator = _build_type_validator(attr)

    return AttributeConstraint(
        namespace=ns,
        local_name=attr.local_name,
        type_validator=type_validator,
        required=attr.required,
        introduced_version=attr.version,
    )


def _build_type_validator(attr: SdkAttribute) -> XsdTypeValidator | None:
    """Build a type validator from SDK attribute metadata.

    Processes the Validators array from SDK JSON to create properly
    constrained validators (enum, string patterns/lengths, number bounds).
    When the SDK provides version-specific validator snapshots, select the
    newest validator set supported by the current file format.
    """
    validators_by_version: dict[str | None, list[dict]] = {}

    for validator in attr.validators:
        if validator.get("Name") in {"RequiredValidator", "OfficeVersionValidator"}:
            continue
        version = validator.get("Version")
        validators_by_version.setdefault(version, []).append(validator)

    if not validators_by_version:
        return _build_default_type_validator(attr)

    branches: list[tuple[str | None, XsdTypeValidator]] = []
    for version, validators in validators_by_version.items():
        branch = _build_validator_for_version(validators, attr)
        if branch is not None:
            branches.append((version, branch))

    if not branches:
        return _build_default_type_validator(attr)
    if len(branches) == 1 and branches[0][0] is None:
        return branches[0][1]
    return VersionedTypeValidator(branches)


def _build_validator_for_version(
    validators: list[dict],
    attr: SdkAttribute,
) -> XsdTypeValidator | None:
    """Build a validator for one SDK version snapshot."""
    union_groups: dict[int, list[dict]] = {}
    simple_validators: list[dict] = []

    for v in validators:
        union_id = v.get("UnionId")
        if union_id is not None:
            union_groups.setdefault(union_id, []).append(v)
        else:
            simple_validators.append(v)

    if union_groups:
        members: list[XsdTypeValidator] = []
        for uid in sorted(union_groups):
            member = _build_validator_from_group(union_groups[uid], attr)
            if member is not None:
                members.append(member)
        if len(members) > 1:
            return UnionTypeValidator(members)
        if len(members) == 1:
            return members[0]

    if simple_validators:
        return _build_validator_from_group(simple_validators, attr)

    return _build_default_type_validator(attr)


def _build_validator_from_group(
    validators: list[dict], attr: SdkAttribute
) -> XsdTypeValidator | None:
    """Build a validator from a group of SDK validators (same UnionId).

    Multiple validators within the same group represent alternatives —
    a value is valid if it matches ANY of them.
    """
    is_hex = attr.type_name == "HexBinaryValue"
    attr_enum_values = _get_enum_values_from_sdk_type(attr.type_name)
    members: list[XsdTypeValidator] = []

    for v in validators:
        name = v.get("Name", "")
        args = _parse_validator_args(v)

        if name == "EnumValidator":
            enum_values = _resolve_enum_values(v, attr)
            if enum_values is not None:
                members.append(StringTypeValidator(enumeration=enum_values))

        elif name == "StringValidator":
            vtype = v.get("Type", "")
            hex_binary = is_hex or "Hex" in vtype
            member = _build_string_validator(args, hex_binary=hex_binary)
            if attr_enum_values is not None:
                member = _apply_enumeration_to_string_validator(member, attr_enum_values)
            members.append(member)

        elif name == "NumberValidator":
            members.append(_build_number_validator(args, attr, v))

    if len(members) > 1:
        return UnionTypeValidator(members)
    if len(members) == 1:
        return members[0]

    return _build_default_type_validator(attr)


def _build_default_type_validator(attr: SdkAttribute) -> XsdTypeValidator | None:
    """Build a validator from the attribute's type name alone."""
    return _build_type_validator_from_sdk_type_name(attr.type_name)


def _build_type_validator_from_sdk_type_name(sdk_type_name: str) -> XsdTypeValidator | None:
    list_item_type = _extract_list_item_type_name(sdk_type_name)
    if list_item_type is not None:
        item_validator = _build_type_validator_from_sdk_type_name(list_item_type)
        return ListTypeValidator(item_validator) if item_validator is not None else None

    enum_values = _get_enum_values_from_sdk_type(sdk_type_name)
    if enum_values is not None:
        return StringTypeValidator(enumeration=enum_values)

    xsd_type = get_xsd_type_name(sdk_type_name)
    return get_type_validator(xsd_type)


def _get_enum_values_from_sdk_type(sdk_type_name: str) -> list[str] | None:
    enum_type = _extract_enum_type_name(sdk_type_name)
    if enum_type is None:
        return None
    return get_enum_values(enum_type)


def _apply_enumeration_to_string_validator(
    validator: XsdTypeValidator,
    enum_values: list[str],
) -> XsdTypeValidator:
    if not isinstance(validator, StringTypeValidator):
        return validator

    pattern = validator.pattern.pattern if validator.pattern is not None else None
    return StringTypeValidator(
        min_length=validator.min_length,
        max_length=validator.max_length,
        pattern=pattern,
        enumeration=enum_values,
    )


def _resolve_enum_values(validator: dict, attr: SdkAttribute) -> list[str] | None:
    validator_type = validator.get("Type")
    if isinstance(validator_type, str):
        enum_values = get_enum_values(validator_type)
        if enum_values is not None:
            return enum_values

    enum_type = _extract_enum_type_name(attr.type_name)
    if enum_type is not None:
        return get_enum_values(enum_type)
    return None


def _parse_validator_args(v: dict) -> dict[str, str]:
    """Extract named arguments from a validator dict."""
    args: dict[str, str] = {}
    for arg in v.get("Arguments", []):
        arg_name = arg.get("Name", "")
        arg_value = arg.get("Value", "")
        if arg_name:
            args[arg_name] = arg_value
    return args


def _build_string_validator(args: dict[str, str], *, hex_binary: bool = False) -> XsdTypeValidator:
    """Build a StringTypeValidator from SDK StringValidator arguments.

    For HexBinaryValue attributes, SDK Length/MinLength/MaxLength refer to
    byte count. Each byte is 2 hex characters, so lengths are doubled.
    """
    min_length = None
    max_length = None
    pattern = None
    scale = 2 if hex_binary else 1

    if "MinLength" in args:
        min_length = int(args["MinLength"]) * scale
    if "MaxLength" in args:
        max_length = int(args["MaxLength"]) * scale
    if "Length" in args:
        length = int(args["Length"]) * scale
        min_length = length
        max_length = length
    if "Pattern" in args:
        pattern = args["Pattern"]

    # Semantic type flags override to specialized validators
    if args.get("IsQName") == "True":
        return QNameTypeValidator()
    if args.get("IsNcName") == "True" or args.get("IsId") == "True":
        return NCNameTypeValidator()
    if args.get("IsUri") == "True":
        return AnyURITypeValidator()

    return StringTypeValidator(
        min_length=min_length,
        max_length=max_length,
        pattern=pattern,
    )


def _build_number_validator(
    args: dict[str, str],
    attr: SdkAttribute,
    validator: dict,
) -> XsdTypeValidator:
    """Build a numeric validator from SDK NumberValidator arguments."""
    base_validator = _get_number_base_validator(attr, validator.get("Type"))
    is_list = _validator_uses_list(validator)

    if isinstance(base_validator, DecimalTypeValidator):
        min_value = base_validator.min_value
        max_value = base_validator.max_value
        min_inclusive = base_validator.min_inclusive
        max_inclusive = base_validator.max_inclusive

        if "MinInclusive" in args:
            min_value = Decimal(args["MinInclusive"])
            min_inclusive = True
        if "MaxInclusive" in args:
            max_value = Decimal(args["MaxInclusive"])
            max_inclusive = True
        if "MinExclusive" in args:
            min_value = Decimal(args["MinExclusive"])
            min_inclusive = False
        if "MaxExclusive" in args:
            max_value = Decimal(args["MaxExclusive"])
            max_inclusive = False
        if args.get("IsPositive") == "True" and min_value is None:
            min_value = Decimal("0")
            min_inclusive = False
        if args.get("IsNonNegative") == "True" and min_value is None:
            min_value = Decimal("0")

        scalar = DecimalTypeValidator(
            min_value=min_value,
            max_value=max_value,
            min_inclusive=min_inclusive,
            max_inclusive=max_inclusive,
        )
        return ListTypeValidator(scalar) if is_list else scalar

    if isinstance(base_validator, IntegerTypeValidator):
        min_value = base_validator.min_value
        max_value = base_validator.max_value
        min_inclusive = base_validator.min_inclusive
        max_inclusive = base_validator.max_inclusive

        if "MinInclusive" in args:
            min_value = int(args["MinInclusive"])
            min_inclusive = True
        if "MaxInclusive" in args:
            max_value = int(args["MaxInclusive"])
            max_inclusive = True
        if "MinExclusive" in args:
            min_value = int(args["MinExclusive"])
            min_inclusive = False
        if "MaxExclusive" in args:
            max_value = int(args["MaxExclusive"])
            max_inclusive = False
        if args.get("IsPositive") == "True" and min_value is None:
            min_value = 0
            min_inclusive = False
        if args.get("IsNonNegative") == "True" and min_value is None:
            min_value = 0

        scalar = IntegerTypeValidator(
            min_value=min_value,
            max_value=max_value,
            min_inclusive=min_inclusive,
            max_inclusive=max_inclusive,
        )
        return ListTypeValidator(scalar) if is_list else scalar

    if base_validator is not None:
        return ListTypeValidator(base_validator) if is_list else base_validator

    fallback = IntegerTypeValidator()
    return ListTypeValidator(fallback) if is_list else fallback


def _get_number_base_validator(
    attr: SdkAttribute,
    validator_type: object,
) -> XsdTypeValidator | None:
    if isinstance(validator_type, str):
        if validator_type.startswith("xsd:"):
            xsd_type = validator_type.split(":", 1)[1]
            built_in = get_type_validator(xsd_type)
            if built_in is not None:
                return built_in

        custom = _CUSTOM_NUMBER_TYPE_VALIDATORS.get(validator_type)
        if custom is not None:
            return custom

    default_validator = _build_default_type_validator(attr)
    if isinstance(default_validator, (DecimalTypeValidator, IntegerTypeValidator)):
        return default_validator
    return None


def _validator_uses_list(validator: dict) -> bool:
    value = validator.get("IsList")
    return value is True or value == "True"


def _convert_particle(
    particle: SdkParticle,
    target_namespace: str,
    namespace_map: dict[str, str],
) -> ParticleConstraint | None:
    """Convert an SDK particle to a ParticleConstraint."""
    kind = particle.kind

    def maybe_collapse_single_child(
        children: list[ParticleConstraint],
    ) -> ParticleConstraint | None:
        if len(children) != 1:
            return None
        if particle.min_occurs != 1 or particle.max_occurs != 1:
            return None
        return children[0]

    if kind == "Sequence":
        children = [
            _convert_particle(item, target_namespace, namespace_map) for item in particle.items
        ]
        children = [c for c in children if c is not None]
        flattened: list[ParticleConstraint] = []
        for child in children:
            if isinstance(child, SequenceParticle) and child.max_occurs == 1:
                optional = child.min_occurs == 0
                if optional and any(sub.min_occurs > 0 for sub in child.children):
                    flattened.append(child)
                    continue
                for sub in child.children:
                    if optional and sub.min_occurs > 0:
                        sub.min_occurs = 0
                    flattened.append(sub)
            else:
                flattened.append(child)
        children = flattened
        collapsed = maybe_collapse_single_child(children)
        if collapsed is not None:
            return collapsed
        return SequenceParticle(
            children=children,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif kind == "Choice":
        children = [
            _convert_particle(item, target_namespace, namespace_map) for item in particle.items
        ]
        children = [c for c in children if c is not None]
        flattened = []
        for child in children:
            if (
                isinstance(child, ChoiceParticle)
                and child.min_occurs == 1
                and child.max_occurs == 1
            ):
                flattened.extend(child.children)
            else:
                flattened.append(child)
        children = flattened
        collapsed = maybe_collapse_single_child(children)
        if collapsed is not None:
            return collapsed
        return ChoiceParticle(
            children=children,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif kind == "All":
        children = [
            _convert_particle(item, target_namespace, namespace_map) for item in particle.items
        ]
        children = [c for c in children if c is not None]
        collapsed = maybe_collapse_single_child(children)
        if collapsed is not None:
            return collapsed
        return AllParticle(
            children=children,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif kind == "Group":
        # Group is a reference to a named group - inline its items
        children = [
            _convert_particle(item, target_namespace, namespace_map) for item in particle.items
        ]
        children = [c for c in children if c is not None]
        if len(children) == 1:
            # Single child group - just return the child
            child = children[0]
            # Apply group's occurrence to child
            if isinstance(child, (SequenceParticle, ChoiceParticle, AllParticle)):
                child.min_occurs = particle.min_occurs
                child.max_occurs = particle.max_occurs
            return child
        else:
            # Multiple children in group - wrap in sequence
            return SequenceParticle(
                children=children,
                min_occurs=particle.min_occurs,
                max_occurs=particle.max_occurs,
            )

    elif kind == "Any":
        return AnyParticle(
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
        )

    elif particle.name:
        # Element reference like "a:CT_OfficeArtExtensionList/a:extLst"
        # Extract the element local name
        if "/" in particle.name:
            elem_ref = particle.name.split("/")[1]
            if ":" in elem_ref:
                prefix, local_name = elem_ref.split(":", 1)
                ns = namespace_map.get(prefix, target_namespace)
            else:
                local_name = elem_ref
                ns = target_namespace
        else:
            local_name = particle.name.split(":")[-1]
            ns = target_namespace

        # Look up version from the referenced element type
        registry = get_registry()
        ref_type = registry.get_type(particle.name)
        introduced_version = ref_type.version if ref_type else None

        return ElementParticle(
            namespace=ns,
            local_name=local_name,
            min_occurs=particle.min_occurs,
            max_occurs=particle.max_occurs,
            introduced_version=introduced_version,
        )

    return None


def _apply_particle_compat_overrides(
    elem_type_name: str,
    content_model: ParticleConstraint | None,
) -> ParticleConstraint | None:
    """Apply targeted schema-model compatibility fixes for SDK parity."""
    if content_model is None:
        return None

    # SDK metadata currently models w:cols child w:col as required (min=1),
    # but SDK validator accepts empty w:cols in equal-width/default scenarios.
    if (
        elem_type_name == "w:CT_Columns/w:cols"
        and isinstance(content_model, SequenceParticle)
        and len(content_model.children) == 1
        and isinstance(content_model.children[0], ElementParticle)
        and content_model.children[0].local_name == "col"
        and content_model.children[0].min_occurs == 1
    ):
        content_model.children[0].min_occurs = 0

    if (
        elem_type_name in {"w:CT_Footnotes/w:footnotes", "w:CT_Endnotes/w:endnotes"}
        and isinstance(content_model, SequenceParticle)
        and len(content_model.children) == 1
        and isinstance(content_model.children[0], ElementParticle)
        and content_model.children[0].max_occurs == 1
    ):
        content_model.children[0].max_occurs = -1

    if (
        elem_type_name == "wps:CT_WordprocessingShape/wps:wsp"
        and isinstance(content_model, SequenceParticle)
        and len(content_model.children) >= 3
        and isinstance(content_model.children[0], ElementParticle)
        and content_model.children[0].local_name == "cNvPr"
        and isinstance(content_model.children[1], ChoiceParticle)
    ):
        c_nv_choice = content_model.children[1]
        if all(
            isinstance(child, ElementParticle)
            and child.local_name in {"cNvSpPr", "cNvCnPr"}
            for child in c_nv_choice.children
        ):
            # Google Docs/Word exports can place the shape/connector non-visual
            # properties before cNvPr. The .NET SDK validator accepts that order
            # for wps:wsp, so keep the SDK metadata sequence and add that
            # app-compatible order as a peer branch.
            reordered = SequenceParticle(
                children=[
                    content_model.children[1],
                    content_model.children[0],
                    *content_model.children[2:],
                ],
                min_occurs=content_model.min_occurs,
                max_occurs=content_model.max_occurs,
            )
            return ChoiceParticle(
                children=[content_model, reordered],
                min_occurs=content_model.min_occurs,
                max_occurs=content_model.max_occurs,
            )

    return content_model


@lru_cache(maxsize=1)
def _build_namespace_map() -> dict[str, str]:
    """Build prefix -> namespace URI map from registry."""
    registry = get_registry()
    registry.load()
    return dict(registry._prefixes)


@lru_cache(maxsize=2048)
def get_element_constraint(tag: str) -> ElementConstraint | None:
    """Get element constraint from SDK schema by Clark notation tag.

    Args:
        tag: Element tag in Clark notation, e.g., "{namespace}localname"

    Returns:
        ElementConstraint if found, None otherwise.
    """
    registry = get_registry()
    elem_type = registry.get_element_type_by_tag(tag)

    if elem_type is None:
        return None

    return convert_element_type(elem_type)


def get_element_constraint_for_element(
    tag: str, element: etree._Element
) -> ElementConstraint | None:
    """Get the best element constraint for a specific element instance."""
    registry = get_registry()
    candidates = registry.get_element_type_candidates(tag)

    if not candidates:
        elem_type = registry.get_element_type_by_tag(tag)
        return convert_element_type(elem_type) if elem_type else None

    if len(candidates) == 1:
        cached = _single_candidate_cache.get(tag)
        if cached is not None:
            return cached
        result = convert_element_type(candidates[0])
        _single_candidate_cache[tag] = result
        return result

    # Multiple candidates: resolution reads only the structural signature
    # below, so memoize on it instead of re-scoring every occurrence (this is
    # the schema hot path on large documents — thousands of repeated shapes
    # collapse to a few hundred distinct signatures).
    key = _multi_candidate_key(tag, element)
    cached_multi = _multi_candidate_cache.get(key, _MULTI_CACHE_MISS)
    if cached_multi is not _MULTI_CACHE_MISS:
        return cached_multi
    elem_type = _get_sdk_element_type_for_element(tag, element)
    result = convert_element_type(elem_type) if elem_type else None
    _multi_candidate_cache[key] = result
    return result


def _multi_candidate_key(
    tag: str, element: etree._Element
) -> tuple[str, str | None, str | None, str | None, tuple[str, ...], tuple[str, ...]]:
    """Signature capturing everything the multi-candidate resolver reads.

    _select_candidate_by_context reads parent namespace plus parent/grandparent local names;
    _score_candidate reads only child tags and child count;
    _missing_required_attributes reads which attributes are present. Child and
    attribute order do not affect the result, so both are sorted to maximise
    cache hits.
    """
    parent = element.getparent()
    parent_ns, parent_local = _split_element_tag(parent)
    grandparent = parent.getparent() if parent is not None else None
    _, grandparent_local = _split_element_tag(grandparent)
    child_tags = tuple(sorted(c.tag for c in element if isinstance(c.tag, str)))
    attr_names = tuple(sorted(element.attrib.keys()))
    return (tag, parent_ns, parent_local, grandparent_local, child_tags, attr_names)


def _get_sdk_element_type_for_element(
    tag: str,
    element: etree._Element,
) -> SdkElementType | None:
    """Resolve the best SDK element type candidate for a specific element instance."""
    registry = get_registry()
    candidates = registry.get_element_type_candidates(tag)

    if not candidates:
        return registry.get_element_type_by_tag(tag)

    if len(candidates) == 1:
        return candidates[0]

    selected_by_context = _select_candidate_by_context(tag, element, candidates)
    if selected_by_context is not None:
        return selected_by_context

    children = [c for c in element if isinstance(c.tag, str)]
    attribute_names = tuple(element.attrib.keys())
    best: SdkElementType | None = None
    best_score: tuple[int, int, int, int, int, int] | None = None

    for candidate in candidates:
        constraint = convert_element_type(candidate)
        if _missing_required_attributes(constraint, element):
            continue
        score = _score_candidate(constraint, children, attribute_names)
        if best_score is None or score > best_score:
            best = candidate
            best_score = score

    if best is not None:
        return best

    return registry.get_element_type_by_tag(tag)


def _select_candidate_by_context(
    tag: str,
    element: etree._Element,
    candidates: list[SdkElementType],
) -> SdkElementType | None:
    parent = element.getparent()
    parent_ns, parent_local = _split_element_tag(parent)
    grandparent = parent.getparent() if parent is not None else None
    _grandparent_ns, grandparent_local = _split_element_tag(grandparent)

    if tag == f"{{{SPREADSHEETML}}}c":
        in_calc_chain = parent_local == "calcChain"
        for candidate in candidates:
            is_calc_candidate = (
                candidate.class_name == "CalculationCell" or "CT_CalcCell" in candidate.name
            )
            if in_calc_chain and is_calc_candidate:
                return candidate
            if not in_calc_chain and not is_calc_candidate:
                return candidate
        return None

    if tag == f"{{{DRAWINGML_CHART}}}ser":
        return _select_candidate_by_class_name(
            candidates,
            _CHART_SERIES_CANDIDATE_BY_PARENT.get(parent_local),
        )

    if tag == f"{{{DRAWINGML_CHART}}}grouping":
        if parent_local in _BAR_CHART_LOCALS:
            return _select_candidate_by_class_name(candidates, "BarGrouping")
        if parent_local and parent_local.endswith("Chart"):
            return _select_candidate_by_class_name(candidates, "Grouping")

    if tag == f"{{{DRAWINGML_CHART}}}order":
        if parent_local == "ser":
            return _select_candidate_by_class_name(candidates, "Order")
        if parent_local == "trendline":
            return _select_candidate_by_class_name(candidates, "PolynomialOrder")

    if tag == f"{{{DRAWINGML_CHART}}}extLst":
        return _select_candidate_by_class_name(
            candidates,
            _CHART_EXTENSION_LIST_CANDIDATE_BY_PARENT.get(parent_local),
        )

    if tag == f"{{{DRAWINGML_CHART}}}ext" and parent_local == "extLst":
        return _select_candidate_by_class_name(
            candidates,
            _CHART_EXTENSION_CANDIDATE_BY_PARENT.get(grandparent_local),
        )

    if tag == f"{{{SPREADSHEETML}}}extLst":
        return _select_candidate_by_class_name(
            candidates,
            _SPREADSHEET_EXTENSION_LIST_CANDIDATE_BY_PARENT.get(parent_local),
        )

    if tag == f"{{{SPREADSHEETML}}}ext" and parent_local == "extLst":
        return _select_candidate_by_class_name(
            candidates,
            _SPREADSHEET_EXTENSION_CANDIDATE_BY_PARENT.get(grandparent_local),
        )

    if tag == f"{{{WORDPROCESSINGML}}}del":
        if parent_ns == OFFICE_DOC_MATH and parent_local == "ctrlPr":
            return _select_candidate_by_class_name(candidates, "DeletedMathControl")
        if any(
            isinstance(child.tag, str) and child.tag == f"{{{WORDPROCESSINGML}}}rPr"
            for child in element
        ):
            return _select_candidate_by_class_name(candidates, "DeletedMathControl")
        if parent_local == "trPr":
            return _select_candidate_by_class_name(candidates, "Deleted")
        return None

    if tag in (f"{{{WORDPROCESSINGML}}}start", f"{{{WORDPROCESSINGML}}}end"):
        expected_type = _WORD_START_END_TYPE_BY_PARENT.get(parent_local)
        if expected_type is not None:
            for candidate in candidates:
                if candidate.type_name == expected_type:
                    return candidate
        return None

    if tag == f"{{{WORDPROCESSINGML}}}rPr":
        return _select_candidate_by_class_name(
            candidates,
            _WORD_RUN_PROPERTIES_CANDIDATE_BY_PARENT.get(parent_local),
        )

    if tag == f"{{{WORDPROCESSINGML}}}pPr":
        return _select_candidate_by_class_name(
            candidates,
            _WORD_PARAGRAPH_PROPERTIES_CANDIDATE_BY_PARENT.get(parent_local),
        )

    return None


def _split_element_tag(element: etree._Element | None) -> tuple[str | None, str | None]:
    if element is None or not isinstance(element.tag, str):
        return None, None
    if element.tag.startswith("{"):
        ns_end = element.tag.index("}")
        return element.tag[1:ns_end], element.tag[ns_end + 1 :]
    return None, element.tag


def _select_candidate_by_class_name(
    candidates: list[SdkElementType],
    class_name: str | None,
) -> SdkElementType | None:
    if class_name is None:
        return None
    for candidate in candidates:
        if candidate.class_name == class_name:
            return candidate
    return None


def convert_element_type(elem_type: SdkElementType) -> ElementConstraint:
    """Convert an SDK element type to an ElementConstraint."""
    return _convert_element_type_by_name(elem_type.name)


@lru_cache(maxsize=8192)
def _convert_element_type_by_name(elem_type_name: str) -> ElementConstraint:
    registry = get_registry()
    elem_type = registry.get_type(elem_type_name)
    if elem_type is None:
        raise ValueError(f"Unknown SDK element type: {elem_type_name}")
    return _convert_element_type_uncached(elem_type)


def _convert_element_type_uncached(elem_type: SdkElementType) -> ElementConstraint:
    """Convert an SDK element type to an ElementConstraint.

    Args:
        elem_type: The SDK element type to convert.

    Returns:
        The converted ElementConstraint.
    """
    registry = get_registry()
    namespace_map = _build_namespace_map()

    # Get schema for target namespace
    schema = None
    for _ns, s in registry._schemas.items():
        if elem_type in s.types:
            schema = s
            break

    target_namespace = schema.target_namespace if schema else ""

    # Convert attributes
    attributes = [_convert_attribute(attr, namespace_map) for attr in elem_type.attributes]

    # Convert particle (content model)
    content_model = None
    if elem_type.particle:
        content_model = _convert_particle(
            elem_type.particle,
            target_namespace,
            namespace_map,
        )
    content_model = _apply_particle_compat_overrides(elem_type.name, content_model)

    # Determine namespace from element name
    if elem_type.element_prefix:
        ns = namespace_map.get(elem_type.element_prefix, target_namespace)
    else:
        ns = target_namespace

    return ElementConstraint(
        namespace=ns,
        local_name=elem_type.element_name or "",
        attributes=attributes,
        content_model=content_model,
        allows_text=False,  # Could be determined from base class
        introduced_version=elem_type.version,
    )


def _missing_required_attributes(
    constraint: ElementConstraint,
    element: etree._Element,
) -> bool:
    for attr in constraint.get_required_attributes():
        if attr.qualified_name not in element.attrib:
            return True
    return False


def _score_candidate(
    constraint: ElementConstraint,
    children: list[etree._Element],
    attribute_names: tuple[str, ...],
) -> tuple[int, int, int, int, int, int]:
    if constraint.content_model is None or not children:
        specific_matches = 0
        total_matches = 0
    else:
        allowed, has_any = _collect_allowed_tags(constraint.content_model)
        specific_matches = sum(1 for child in children if child.tag in allowed)
        total_matches = len(children) if has_any else specific_matches

    declared_attrs = {attr.qualified_name for attr in constraint.attributes}
    matched_attrs = 0
    for attr_name in attribute_names:
        if attr_name in declared_attrs:
            matched_attrs += 1
    undeclared_attrs = len(attribute_names) - matched_attrs

    if children:
        shape_fit = 1 if constraint.content_model is not None else 0
    else:
        shape_fit = 1 if constraint.content_model is None else 0

    return (
        specific_matches,
        total_matches,
        shape_fit,
        matched_attrs,
        -undeclared_attrs,
        len(constraint.attributes),
    )


def _collect_allowed_tags(
    particle: ParticleConstraint,
) -> tuple[set[str], bool]:
    allowed: set[str] = set()
    has_any = False

    def visit(node: ParticleConstraint) -> None:
        nonlocal has_any
        if isinstance(node, ElementParticle):
            allowed.add(node.qualified_name)
        elif isinstance(node, AnyParticle):
            has_any = True
        elif isinstance(node, CompositeParticle):
            for child in node.children:
                visit(child)

    visit(particle)
    return allowed, has_any


def get_sdk_element_info(tag: str) -> dict | None:
    """Get raw SDK element info for debugging/inspection.

    Args:
        tag: Element tag in Clark notation.

    Returns:
        Dictionary with element info, or None.
    """
    registry = get_registry()
    elem_type = registry.get_element_type_by_tag(tag)

    if elem_type is None:
        return None

    return {
        "name": elem_type.name,
        "class_name": elem_type.class_name,
        "base_class": elem_type.base_class,
        "is_abstract": elem_type.is_abstract,
        "is_leaf": elem_type.is_leaf_element,
        "attributes": [
            {
                "qname": a.qname,
                "type": a.type_name,
                "required": a.required,
            }
            for a in elem_type.attributes
        ],
        "has_particle": elem_type.particle is not None,
    }
