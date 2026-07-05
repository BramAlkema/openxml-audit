"""Whole-dataset bridge invariants for shipped Open XML SDK schema data."""

from __future__ import annotations

from collections import Counter

from lxml import etree

from openxml_audit.codegen.constraint_bridge import _build_type_validator
from openxml_audit.context import ValidationContext
from openxml_audit.errors import FileFormat
from openxml_audit.schema.types import (
    AnyURITypeValidator,
    DecimalTypeValidator,
    IntegerTypeValidator,
    NCNameTypeValidator,
    QNameTypeValidator,
    StringTypeValidator,
    UnionTypeValidator,
    VersionedTypeValidator,
)
from tests.bridge_invariant_helpers import (
    args_dict,
    build_number_validator_for_use,
    build_string_validator_for_use,
    build_versioned_validator_for_use,
    collect_ambiguous_element_candidates,
    collect_live_enum_attribute_uses,
    collect_live_number_validator_uses,
    collect_live_string_validator_uses,
    collect_live_versioned_validator_uses,
    expected_selected_branch_for_format,
    find_fractional_decimal_probe,
    find_transition_probe,
    find_valid_numeric_probe,
    get_attribute_for_use,
    get_lower_bound_rejection_probe,
    get_selected_branch_version_for_format,
    get_selected_candidate_class_name,
    get_upper_bound_rejection_probe,
    partition_live_number_validator_uses,
    runtime_has_specialized_string_validator,
    runtime_has_string_pattern,
    unwrap_list_item_validator,
    validator_has_enumeration_semantics,
    versioned_probe_matrix,
)

_TEST_NSMAP = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mso14": "http://schemas.microsoft.com/office/2009/07/customui",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}
_CHART_NS = _TEST_NSMAP["c"]
_CUSTOMUI_NS = _TEST_NSMAP["mso14"]
_MATH_NS = _TEST_NSMAP["m"]
_SPREADSHEET_NS = _TEST_NSMAP["x"]
_WORD_NS = _TEST_NSMAP["w"]


def _select_ambiguous_case(
    xml: bytes,
    xpath: str = ".",
) -> tuple[str, str | None, tuple[str, ...]]:
    root = etree.fromstring(xml)
    if xpath == ".":
        element = root
    else:
        result = root.xpath(xpath, namespaces=_TEST_NSMAP)
        assert len(result) == 1
        element = result[0]

    assert isinstance(element, etree._Element)
    tag = element.tag
    assert isinstance(tag, str)
    return (
        tag,
        get_selected_candidate_class_name(element),
        collect_ambiguous_element_candidates().get(tag, ()),
    )


def test_shipped_enum_attributes_do_not_degrade_to_plain_strings() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_enum_attribute_uses():
        if not use.enum_type_names:
            continue

        checked += 1
        validator = _build_type_validator(get_attribute_for_use(use))
        if validator_has_enumeration_semantics(validator):
            continue

        validator_name = type(validator).__name__ if validator is not None else "None"
        detail = ""
        if isinstance(validator, StringTypeValidator):
            detail = " plain StringTypeValidator"
        failures.append(f"{use.owner} ({use.sdk_type_name}) -> {validator_name}{detail}")

    assert checked > 0
    assert failures == []


def test_wrapped_union_enum_validators_preserve_enumeration_semantics() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_enum_attribute_uses():
        if (
            not use.has_enum_validator
            or not use.expects_union_wrapper
            or use.expects_versioned_wrapper
        ):
            continue

        checked += 1
        validator = _build_type_validator(get_attribute_for_use(use))
        if isinstance(validator, UnionTypeValidator) and validator_has_enumeration_semantics(
            validator
        ):
            continue

        validator_name = type(validator).__name__ if validator is not None else "None"
        failures.append(f"{use.owner} ({use.sdk_type_name}) -> {validator_name}")

    assert checked > 0
    assert failures == []


def test_wrapped_versioned_enum_validators_preserve_enumeration_semantics() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_enum_attribute_uses():
        if not use.has_enum_validator or not use.expects_versioned_wrapper:
            continue

        checked += 1
        validator = _build_type_validator(get_attribute_for_use(use))
        if isinstance(validator, VersionedTypeValidator) and validator_has_enumeration_semantics(
            validator
        ):
            continue

        validator_name = type(validator).__name__ if validator is not None else "None"
        failures.append(f"{use.owner} ({use.sdk_type_name}) -> {validator_name}")

    assert checked > 0
    assert failures == []


def test_number_validator_dataset_partitions_cover_integer_decimal_bounds_and_lists() -> None:
    partitions = partition_live_number_validator_uses()

    assert partitions["integer"] > 0
    assert partitions["decimal"] > 0
    assert partitions["min_bounded"] > 0
    assert partitions["max_bounded"] > 0
    assert partitions["list"] > 0


def test_nonnegative_number_validator_branches_reject_negative_probes() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_number_validator_uses():
        validator = build_number_validator_for_use(use)
        item_validator = unwrap_list_item_validator(validator)
        if not isinstance(item_validator, (IntegerTypeValidator, DecimalTypeValidator)):
            continue
        if item_validator.min_value is None or item_validator.min_value < 0:
            continue

        checked += 1
        probe = "-1"
        if validator.validate(probe).is_valid:
            failures.append(f"{use.owner} ({use.sdk_type_name}) accepted probe {probe}")

    assert checked > 0
    assert failures == []


def test_decimal_number_validator_branches_accept_fractional_probes() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_number_validator_uses():
        validator = build_number_validator_for_use(use)
        item_validator = unwrap_list_item_validator(validator)
        if not isinstance(item_validator, DecimalTypeValidator):
            continue

        checked += 1
        probe = find_fractional_decimal_probe(item_validator)
        if probe is None:
            failures.append(f"{use.owner} ({use.sdk_type_name}) had no valid fractional probe")
            continue
        if not validator.validate(probe).is_valid:
            failures.append(f"{use.owner} ({use.sdk_type_name}) rejected probe {probe}")

    assert checked > 0
    assert failures == []


def test_list_valued_number_validator_branches_preserve_list_semantics() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_number_validator_uses():
        if not use.is_list:
            continue

        checked += 1
        validator = build_number_validator_for_use(use)
        item_validator = unwrap_list_item_validator(validator)
        if not isinstance(item_validator, (IntegerTypeValidator, DecimalTypeValidator)):
            failures.append(f"{use.owner} ({use.sdk_type_name}) built non-numeric list validator")
            continue

        scalar_probe = find_valid_numeric_probe(item_validator)
        if scalar_probe is None:
            failures.append(f"{use.owner} ({use.sdk_type_name}) had no valid scalar probe")
            continue

        valid_list_probe = f"{scalar_probe} {scalar_probe}"
        invalid_list_probe = f"{scalar_probe} not-a-number"
        if not validator.validate(valid_list_probe).is_valid:
            failures.append(f"{use.owner} ({use.sdk_type_name}) rejected probe {valid_list_probe}")
        if validator.validate(invalid_list_probe).is_valid:
            failures.append(
                f"{use.owner} ({use.sdk_type_name}) accepted probe {invalid_list_probe}"
            )

    assert checked > 0
    assert failures == []


def test_number_validator_range_bounds_survive_bridge_conversion() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_number_validator_uses():
        validator = build_number_validator_for_use(use)
        item_validator = unwrap_list_item_validator(validator)
        if not isinstance(item_validator, (IntegerTypeValidator, DecimalTypeValidator)):
            continue

        lower_probe = get_lower_bound_rejection_probe(item_validator)
        if lower_probe is not None:
            checked += 1
            if validator.validate(lower_probe).is_valid:
                failures.append(
                    f"{use.owner} ({use.sdk_type_name}) accepted lower-bound probe {lower_probe}"
                )

        upper_probe = get_upper_bound_rejection_probe(item_validator)
        if upper_probe is not None:
            checked += 1
            if validator.validate(upper_probe).is_valid:
                failures.append(
                    f"{use.owner} ({use.sdk_type_name}) accepted upper-bound probe {upper_probe}"
                )

    assert checked > 0
    assert failures == []


def test_qname_string_validator_metadata_preserves_qname_semantics() -> None:
    failures: list[str] = []
    checked = 0
    root = etree.fromstring(b'<t:root xmlns:t="urn:test" xmlns:a="urn:a"/>')
    context = ValidationContext()
    context.push_element(root)
    try:
        for use in collect_live_string_validator_uses():
            args = args_dict(use)
            if args.get("IsQName") != "True":
                continue

            checked += 1
            validator = build_string_validator_for_use(use)
            if not runtime_has_specialized_string_validator(validator, QNameTypeValidator):
                failures.append(f"{use.owner} ({use.sdk_type_name}) lost QName specialization")
                continue
            if validator is None:
                failures.append(f"{use.owner} ({use.sdk_type_name}) built no validator")
                continue
            if validator.validate("missing:l", context).is_valid:
                failures.append(f"{use.owner} ({use.sdk_type_name}) accepted probe missing:l")
    finally:
        context.pop_element()

    assert checked > 0
    assert failures == []


def test_ncname_and_uri_string_validator_metadata_preserve_specialized_behavior() -> None:
    failures: list[str] = []
    checked = 0
    context = ValidationContext()

    for use in collect_live_string_validator_uses():
        args = args_dict(use)
        validator = build_string_validator_for_use(use)
        if validator is None:
            continue

        if args.get("IsNcName") == "True" or args.get("IsId") == "True":
            checked += 1
            if not runtime_has_specialized_string_validator(validator, NCNameTypeValidator):
                failures.append(f"{use.owner} ({use.sdk_type_name}) lost NCName specialization")
            elif validator.validate("bad:name", context).is_valid:
                failures.append(f"{use.owner} ({use.sdk_type_name}) accepted probe bad:name")

        if args.get("IsUri") == "True":
            checked += 1
            if not runtime_has_specialized_string_validator(validator, AnyURITypeValidator):
                failures.append(f"{use.owner} ({use.sdk_type_name}) lost anyURI specialization")
            elif validator.validate("<bad>", context).is_valid:
                failures.append(f"{use.owner} ({use.sdk_type_name}) accepted probe <bad>")

    assert checked > 0
    assert failures == []


def test_pattern_only_string_validators_do_not_degrade_to_unconstrained_strings() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_string_validator_uses():
        args = args_dict(use)
        pattern = args.get("Pattern")
        if pattern is None:
            continue
        if any(args.get(flag) == "True" for flag in ("IsQName", "IsNcName", "IsId", "IsUri")):
            continue

        checked += 1
        validator = build_string_validator_for_use(use)
        if not runtime_has_string_pattern(validator, pattern):
            failures.append(f"{use.owner} ({use.sdk_type_name}) lost pattern {pattern!r}")

    assert checked > 0
    assert failures == []


def test_versioned_validator_scanner_covers_two_and_three_branch_families() -> None:
    uses = collect_live_versioned_validator_uses()
    branch_counts = Counter(use.branch_versions for use in uses)

    assert branch_counts[("Office2007", "Office2010")] > 0
    assert branch_counts[("Office2007", "Office2010", "Office2013")] > 0


def test_versioned_validators_select_expected_branch_for_office_2007() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_versioned_validator_uses():
        validator = build_versioned_validator_for_use(use)
        probe_matrix = versioned_probe_matrix(validator)
        if not any(len(set(results)) > 1 for results in probe_matrix.values()):
            continue

        for probe, _results in probe_matrix.items():
            expected_version, expected_branch = expected_selected_branch_for_format(
                validator,
                FileFormat.OFFICE_2007,
            )
            expected = expected_branch.validate(
                probe,
                ValidationContext(file_format=FileFormat.OFFICE_2007),
            ).is_valid
            actual = validator.validate(
                probe,
                ValidationContext(file_format=FileFormat.OFFICE_2007),
            ).is_valid
            checked += 1
            if actual != expected:
                failures.append(
                    f"{use.owner} OFFICE_2007 probe {probe!r} expected {expected} "
                    f"from {expected_version} but got {actual}"
                )

    assert checked > 0
    assert failures == []


def test_versioned_validators_select_expected_branch_for_newer_formats() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_versioned_validator_uses():
        validator = build_versioned_validator_for_use(use)
        for file_format in (FileFormat.OFFICE_2010, FileFormat.OFFICE_2013, FileFormat.OFFICE_2019):
            expected_version, expected_branch = expected_selected_branch_for_format(
                validator,
                file_format,
            )
            actual_version = get_selected_branch_version_for_format(validator, file_format)
            checked += 1
            if actual_version != expected_version:
                failures.append(
                    f"{use.owner} {file_format.name} expected branch {expected_version} "
                    f"got {actual_version}"
                )
                continue

            for probe, _results in versioned_probe_matrix(validator).items():
                expected = expected_branch.validate(
                    probe,
                    ValidationContext(file_format=file_format),
                ).is_valid
                actual = validator.validate(
                    probe,
                    ValidationContext(file_format=file_format),
                ).is_valid
                if actual != expected:
                    failures.append(
                        f"{use.owner} {file_format.name} probe {probe!r} expected {expected} "
                        f"from {expected_version} but got {actual}"
                    )
                    break

    assert checked > 0
    assert failures == []


def test_versioned_validator_branch_evolution_is_not_flattened_into_global_union() -> None:
    failures: list[str] = []
    checked = 0

    for use in collect_live_versioned_validator_uses():
        validator = build_versioned_validator_for_use(use)
        probe = find_transition_probe(
            validator,
            FileFormat.OFFICE_2007,
            FileFormat.OFFICE_2010,
        )
        if probe is None:
            continue

        checked += 1
        older_result = validator.validate(
            probe,
            ValidationContext(file_format=FileFormat.OFFICE_2007),
        ).is_valid
        newer_result = validator.validate(
            probe,
            ValidationContext(file_format=FileFormat.OFFICE_2010),
        ).is_valid
        if older_result or not newer_result:
            failures.append(
                f"{use.owner} probe {probe!r} had OFFICE_2007={older_result} "
                f"OFFICE_2010={newer_result}"
            )

    assert checked > 0
    assert failures == []


def test_ambiguous_element_scanner_covers_phase5_target_families() -> None:
    ambiguous = collect_ambiguous_element_candidates()

    assert "{http://schemas.microsoft.com/office/2009/07/customui}group" in ambiguous
    assert "{http://schemas.openxmlformats.org/drawingml/2006/chart}ser" in ambiguous
    assert "{http://schemas.openxmlformats.org/drawingml/2006/chart}ext" in ambiguous
    assert "{http://schemas.openxmlformats.org/drawingml/2006/chart}extLst" in ambiguous
    assert "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}ext" in ambiguous
    assert "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}extLst" in ambiguous
    assert "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del" in ambiguous
    assert "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr" in ambiguous
    assert "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr" in ambiguous


def test_multi_candidate_resolution_is_cached_by_context_signature() -> None:
    from openxml_audit.codegen import constraint_bridge as cb

    chart_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    ser = f"{{{chart_ns}}}ser"
    assert ser in collect_ambiguous_element_candidates()

    def make_ser(parent_local: str) -> etree._Element:
        root = etree.fromstring(
            f'<c:{parent_local} xmlns:c="{chart_ns}"><c:ser/></c:{parent_local}>'.encode()
        )
        return root[0]

    cb._multi_candidate_cache.clear()

    bar = cb.get_element_constraint_for_element(ser, make_ser("barChart"))
    assert len(cb._multi_candidate_cache) == 1

    # Same signature (different lxml object) -> cache hit, identical result,
    # no new entry.
    bar_again = cb.get_element_constraint_for_element(ser, make_ser("barChart"))
    assert bar_again is bar
    assert len(cb._multi_candidate_cache) == 1

    # Different parent -> different signature -> separate entry resolving to a
    # different candidate. Proves the key distinguishes contexts the resolver
    # treats differently (parent-driven chart-series disambiguation).
    line = cb.get_element_constraint_for_element(ser, make_ser("lineChart"))
    assert len(cb._multi_candidate_cache) == 2
    assert line is not bar

    # The cache distinguishes exactly the candidates the uncached resolver
    # would pick for each parent.
    assert get_selected_candidate_class_name(make_ser("barChart")) == "BarChartSeries"
    assert get_selected_candidate_class_name(make_ser("lineChart")) == "LineChartSeries"


def test_multi_candidate_cache_distinguishes_parent_namespace() -> None:
    from openxml_audit.codegen import constraint_bridge as cb

    deleted = f"{{{_WORD_NS}}}del"
    assert deleted in collect_ambiguous_element_candidates()

    def make_deleted(parent_prefix: str, parent_ns: str) -> etree._Element:
        root = etree.fromstring(
            (
                f'<{parent_prefix}:ctrlPr xmlns:{parent_prefix}="{parent_ns}" '
                f'xmlns:w="{_WORD_NS}"><w:del/></{parent_prefix}:ctrlPr>'
            ).encode()
        )
        return root[0]

    cb._multi_candidate_cache.clear()

    math_ctrl = make_deleted("m", _MATH_NS)
    math_deleted = cb.get_element_constraint_for_element(deleted, math_ctrl)
    assert len(cb._multi_candidate_cache) == 1

    word_ctrl = make_deleted("p", _WORD_NS)
    word_deleted = cb.get_element_constraint_for_element(deleted, word_ctrl)
    assert len(cb._multi_candidate_cache) == 2
    assert word_deleted is not math_deleted

    assert get_selected_candidate_class_name(math_ctrl) == "DeletedMathControl"
    assert get_selected_candidate_class_name(word_ctrl) == "Deleted"


def test_customui_ambiguous_elements_choose_attribute_compatible_candidates() -> None:
    failures: list[str] = []

    cases = [
        (
            (
                f'<mso14:group xmlns:mso14="{_CUSTOMUI_NS}" helperText="help">'
                "<mso14:primaryItem/>"
                "</mso14:group>"
            ),
            ".",
            "BackstageGroup",
        ),
        (
            (
                f'<mso14:group xmlns:mso14="{_CUSTOMUI_NS}" imageMso="HappyFace">'
                '<mso14:button idMso="Copy"/>'
                "</mso14:group>"
            ),
            ".",
            "Group",
        ),
        (
            f'<mso14:button xmlns:mso14="{_CUSTOMUI_NS}" expand="true" style="warning"/>',
            ".",
            "BackstageGroupButton",
        ),
        (
            f'<mso14:control xmlns:mso14="{_CUSTOMUI_NS}" id="qat-copy"/>',
            ".",
            "ControlCloneQat",
        ),
        (
            f'<mso14:menu xmlns:mso14="{_CUSTOMUI_NS}" size="large"/>',
            ".",
            "Menu",
        ),
    ]

    for xml, xpath, expected in cases:
        tag, chosen, candidates = _select_ambiguous_case(xml, xpath)
        if chosen != expected:
            failures.append(f"{tag} expected {expected} got {chosen}; candidates={candidates}")

    assert failures == []


def test_chart_and_spreadsheet_ambiguous_elements_choose_context_appropriate_candidates() -> None:
    failures: list[str] = []

    cases = [
        (
            (
                f'<c:barChart xmlns:c="{_CHART_NS}">'
                '<c:grouping val="clustered"/>'
                '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
                "</c:barChart>"
            ),
            "./c:grouping",
            "BarGrouping",
        ),
        (
            (
                f'<c:barChart xmlns:c="{_CHART_NS}">'
                '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
                "</c:barChart>"
            ),
            "./c:ser",
            "BarChartSeries",
        ),
        (
            (
                f'<c:lineChart xmlns:c="{_CHART_NS}">'
                '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
                '<c:trendline><c:order val="2"/></c:trendline>'
                "</c:lineChart>"
            ),
            "./c:ser",
            "LineChartSeries",
        ),
        (
            (
                f'<c:lineChart xmlns:c="{_CHART_NS}">'
                '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
                '<c:trendline><c:order val="2"/></c:trendline>'
                "</c:lineChart>"
            ),
            "./c:trendline/c:order",
            "PolynomialOrder",
        ),
        (
            f'<c:lineChart xmlns:c="{_CHART_NS}"><c:extLst/></c:lineChart>',
            "./c:extLst",
            "LineChartExtensionList",
        ),
        (
            f'<c:stockChart xmlns:c="{_CHART_NS}"><c:extLst/></c:stockChart>',
            "./c:extLst",
            "StockChartExtensionList",
        ),
        (
            f'<c:lineChart xmlns:c="{_CHART_NS}"><c:extLst><c:ext/></c:extLst></c:lineChart>',
            "./c:extLst/c:ext",
            "LineChartExtension",
        ),
        (
            f'<c:stockChart xmlns:c="{_CHART_NS}"><c:extLst><c:ext/></c:extLst></c:stockChart>',
            "./c:extLst/c:ext",
            "StockChartExtension",
        ),
        (
            f'<x:queryTable xmlns:x="{_SPREADSHEET_NS}"><x:extLst/></x:queryTable>',
            "./x:extLst",
            "QueryTableExtensionList",
        ),
        (
            f'<x:worksheet xmlns:x="{_SPREADSHEET_NS}"><x:extLst/></x:worksheet>',
            "./x:extLst",
            "WorksheetExtensionList",
        ),
        (
            (
                f'<x:queryTable xmlns:x="{_SPREADSHEET_NS}">'
                "<x:extLst><x:ext/></x:extLst>"
                "</x:queryTable>"
            ),
            "./x:extLst/x:ext",
            "QueryTableExtension",
        ),
        (
            (
                f'<x:worksheet xmlns:x="{_SPREADSHEET_NS}">'
                "<x:extLst><x:ext/></x:extLst>"
                "</x:worksheet>"
            ),
            "./x:extLst/x:ext",
            "WorksheetExtension",
        ),
    ]

    for xml, xpath, expected in cases:
        tag, chosen, candidates = _select_ambiguous_case(xml, xpath)
        if chosen != expected:
            failures.append(f"{tag} expected {expected} got {chosen}; candidates={candidates}")

    assert failures == []


def test_wordprocessing_overloaded_tags_choose_context_appropriate_candidates() -> None:
    failures: list[str] = []

    cases = [
        (
            f'<w:style xmlns:w="{_WORD_NS}"><w:pPr/></w:style>',
            "./w:pPr",
            "StyleParagraphProperties",
        ),
        (
            f'<w:lvl xmlns:w="{_WORD_NS}"><w:pPr/></w:lvl>',
            "./w:pPr",
            "PreviousParagraphProperties",
        ),
        (
            f'<w:pPrChange xmlns:w="{_WORD_NS}"><w:pPr/></w:pPrChange>',
            "./w:pPr",
            "ParagraphPropertiesExtended",
        ),
        (
            f'<w:style xmlns:w="{_WORD_NS}"><w:rPr/></w:style>',
            "./w:rPr",
            "StyleRunProperties",
        ),
        (
            f'<w:pPr xmlns:w="{_WORD_NS}"><w:rPr/></w:pPr>',
            "./w:rPr",
            "ParagraphMarkRunProperties",
        ),
        (
            f'<w:rPrChange xmlns:w="{_WORD_NS}"><w:rPr/></w:rPrChange>',
            "./w:rPr",
            "PreviousParagraphMarkRunProperties",
        ),
        (
            (
                f'<m:ctrlPr xmlns:m="{_MATH_NS}" xmlns:w="{_WORD_NS}">'
                "<w:del><w:rPr/></w:del>"
                "</m:ctrlPr>"
            ),
            "./w:del",
            "DeletedMathControl",
        ),
        (
            f'<w:trPr xmlns:w="{_WORD_NS}"><w:del/></w:trPr>',
            "./w:del",
            "Deleted",
        ),
    ]

    for xml, xpath, expected in cases:
        tag, chosen, candidates = _select_ambiguous_case(xml, xpath)
        if chosen != expected:
            failures.append(f"{tag} expected {expected} got {chosen}; candidates={candidates}")

    assert failures == []
