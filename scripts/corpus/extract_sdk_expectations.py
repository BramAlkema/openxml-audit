#!/usr/bin/env python3
"""Extract expected validation outcomes from Open XML SDK tests.

This script parses SDK test sources and updates a corpus manifest with
expectation entries keyed by `source_relpath`.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DEFAULT_SDK_ROOT = Path("/tmp/openxml-sdk-upstream")
DEFAULT_MANIFEST = Path("data/corpus/sdk_seed/manifest.json")
ASSETS_DIR = Path("test/DocumentFormat.OpenXml.Tests.Assets")
TESTS_DIR = Path("test/DocumentFormat.OpenXml.Tests")

CLASS_RE = re.compile(
    r"^\s*public\s+(?:static\s+)?(?:partial\s+)?class\s+([A-Za-z0-9_]+)\s*(?:[:{]|$)"
)
CONST_RE = re.compile(r'^\s*public const string ([A-Za-z0-9_]+)\s*=\s*"([^"]+)";')
METHOD_RE = re.compile(
    r"^\s*public\s+(?:static\s+)?(?:async\s+)?(?:void|Task(?:<[^>]+>)?)\s+([A-Za-z0-9_]+)\s*\("
)
GET_STREAM_RE = re.compile(r"GetStream\(\s*([A-Za-z0-9_\.]+)\s*(?:,|\))")
VALIDATOR_VERSION_RE = re.compile(
    r"new\s+OpenXmlValidator\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*\)"
)
VALIDATOR_DEFAULT_RE = re.compile(r"new\s+OpenXmlValidator\(\s*\)")
INLINE_FILE_COUNT_RE = re.compile(r"\[InlineData\(\s*([A-Za-z0-9_\.]+)\s*,\s*(-?\d+)\s*\)\]")
INLINE_VERSION_COUNT_RE = re.compile(
    r"\[InlineData\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*,\s*(-?\d+)\s*\)\]"
)
INLINE_FILE_ONLY_RE = re.compile(r"\[InlineData\(\s*([A-Za-z0-9_\.]+)\s*\)\]")
INLINE_VERSION_ONLY_RE = re.compile(r"\[InlineData\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*\)\]")
INLINE_FIRST_ARG_RE = re.compile(r"\[InlineData\(\s*([A-Za-z0-9_\.]+)")
VALIDATE_CALL_RE = re.compile(r"\.Validate\s*\(")
PACKAGE_OPEN_RE = re.compile(
    r"\b(?:var|[A-Za-z_][A-Za-z0-9_<>,\.\?]*)\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*"
    r"(?:[A-Za-z_][A-Za-z0-9_\.]*Document|OpenXmlPackage)\.(?:Open|Create)\s*\("
)
MAX_ERRORS_SIGNAL_RE = re.compile(r"\bMaxNumberOfErrors\b")
ASSERT_TRUE_ALLOWED_COUNTS_RE = re.compile(
    r"Assert\.True\(\s*cnt\s*==\s*(-?\d+)\s*\|\|\s*cnt\s*==\s*(-?\d+)\s*\)"
)
XLSX_HELPER_CALL_RE = re.compile(
    r"XlsxValidationHelper\(\s*stream\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)"
)
VALIDATOR_ALIAS_VERSION_RE = re.compile(
    r"\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*new\s+OpenXmlValidator\(\s*FileFormatVersions\.([A-Za-z0-9_]+)\s*\)"
)
VALIDATOR_ALIAS_DEFAULT_RE = re.compile(
    r"\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*new\s+OpenXmlValidator\(\s*\)"
)
ASSIGN_VALIDATE_RE = re.compile(
    r"\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*\.Validate\s*\("
)
ASSERT_EMPTY_VAR_RE = re.compile(r"Assert\.Empty\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)")
ASSERT_SINGLE_VAR_RE = re.compile(r"Assert\.Single\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)")
ASSERT_EQUAL_VAR_COUNT_RE = re.compile(
    r"Assert\.Equal\(\s*(-?\d+)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\.Count\(\)\s*\)"
)
ASSERT_NOTNULL_VAR_RE = re.compile(r"Assert\.NotNull\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)")
ASSERT_EMPTY_DIRECT_RE = re.compile(
    r"Assert\.Empty\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\.Validate\s*\("
)
ASSERT_SINGLE_DIRECT_RE = re.compile(
    r"Assert\.Single\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\.Validate\s*\("
)
ASSERT_NOTNULL_DIRECT_RE = re.compile(
    r"Assert\.NotNull\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\.Validate\s*\("
)
ASSERT_EQUAL_DIRECT_COUNT_RE = re.compile(
    r"Assert\.Equal\(\s*(-?\d+)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\.Validate\s*\("
)
MUTATION_NAME_RE = re.compile(
    r"(Add|Remove|Delete|Insert|Replace|Clear|Change|Create|Set|Update|Modify|Annotation)"
)
MUTATION_SIGNAL_RE = re.compile(
    r"(\.DeletePart\(|\.Add(?:New)?Part\(|\.AddImagePart\(|\.AddPart\(|\.RemovePart\(|"
    r"\.RemoveAllChildren\(|\.Replace(?:Child|All)\(|\.Insert(?:Before|After|At|BeforeSelf|AfterSelf)\(|"
    r"\.SetAttributes?\(|\.RemoveAttributes?\(|\.ClearAllAttributes\(|"
    r"\.InnerXml\s*=|\.InnerText\s*=|\.OuterXml\s*=|"
    r"\bOpen\s*\([^)]*,\s*true\s*\)|\.Save\s*\(|"
    r"\bMarkupCompatibilityProcessSettings\b|\bOpenSettings\b)",
    re.IGNORECASE,
)
VALID_VERSIONS = {"Office2007", "Office2010", "Office2013", "Office2016", "Office2019"}
VALIDATOR_NAME_VERSION_MAP = {
    "12": "Office2007",
    "14": "Office2010",
    "15": "Office2013",
    "16": "Office2016",
    "19": "Office2019",
}


@dataclass(frozen=True)
class MethodBlock:
    """Parsed C# method block."""

    name: str
    start_line: int
    attributes: str
    body: str
    source: str


@dataclass(frozen=True)
class ExtractedExpectation:
    """Expectation mapped to a corpus file."""

    relpath: str
    kind: str
    validator_versions: tuple[str, ...]
    scenario: str
    expected_error_count: int | None
    expected_error_counts: tuple[int, ...]
    source_file: str
    source_method: str
    source_line: int


def _load_manifest(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_manifest(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _build_relpath_indexes(manifest: dict[str, Any]) -> tuple[dict[str, str], dict[str, str]]:
    exact: dict[str, str] = {}
    lower: dict[str, str] = {}
    for entry in manifest.get("files", []):
        relpath = entry.get("source_relpath")
        if not isinstance(relpath, str):
            continue
        dotted = relpath.replace("/", ".")
        exact[dotted] = relpath
        lower[dotted.lower()] = relpath
    return exact, lower


def _resolve_resource_to_relpath(
    resource: str, dotted_exact: dict[str, str], dotted_lower: dict[str, str]
) -> str | None:
    if resource in dotted_exact:
        return dotted_exact[resource]
    lowered = resource.lower()
    if lowered in dotted_lower:
        return dotted_lower[lowered]

    parts = resource.split(".")
    if len(parts) < 2:
        return None

    for split_index in range(len(parts) - 1, 0, -1):
        directory = "/".join(parts[:split_index])
        filename = ".".join(parts[split_index:])
        candidate = f"{directory}/{filename}"
        dotted = candidate.replace("/", ".")
        if dotted in dotted_exact:
            return dotted_exact[dotted]
        if dotted.lower() in dotted_lower:
            return dotted_lower[dotted.lower()]
    return None


def _parse_constant_symbols(assets_dir: Path) -> dict[str, str]:
    symbol_to_resource: dict[str, str] = {}
    unique_short: dict[str, str] = {}
    ambiguous_short: set[str] = set()

    for path in sorted(assets_dir.glob("TestAssets*.cs")):
        lines = path.read_text(encoding="utf-8").splitlines()
        scope: list[str] = []
        pending_classes: list[str] = []

        for line in lines:
            class_match = CLASS_RE.match(line)
            if class_match is not None:
                pending_classes.append(class_match.group(1))

            const_match = CONST_RE.match(line)
            if const_match is not None:
                const_name, resource = const_match.group(1), const_match.group(2)
                if "TestAssets" in scope:
                    idx = scope.index("TestAssets")
                    prefix = scope[idx + 1 :]
                else:
                    prefix = scope[:]
                symbol = ".".join([*prefix, const_name]) if prefix else const_name
                symbol_to_resource[symbol] = resource

                if const_name in ambiguous_short:
                    pass
                elif const_name in unique_short and unique_short[const_name] != resource:
                    ambiguous_short.add(const_name)
                    unique_short.pop(const_name, None)
                else:
                    unique_short[const_name] = resource

            opens = line.count("{")
            closes = line.count("}")
            while opens > 0 and pending_classes:
                scope.append(pending_classes.pop(0))
                opens -= 1
            for _ in range(closes):
                if scope:
                    scope.pop()

    for short_name, resource in unique_short.items():
        symbol_to_resource.setdefault(short_name, resource)
    return symbol_to_resource


def _find_method_end(lines: list[str], start_index: int) -> int:
    open_line = -1
    open_col = -1
    for i in range(start_index, len(lines)):
        col = lines[i].find("{")
        if col >= 0:
            open_line = i
            open_col = col
            break
    if open_line < 0:
        return start_index

    depth = 0
    for i in range(open_line, len(lines)):
        start_col = open_col if i == open_line else 0
        line = lines[i]
        for col in range(start_col, len(line)):
            char = line[col]
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    return i
    return len(lines) - 1


def _collect_attributes(lines: list[str], method_start: int) -> str:
    attrs: list[str] = []
    i = method_start - 1
    while i >= 0:
        stripped = lines[i].strip()
        if stripped.startswith("["):
            attrs.append(lines[i])
            i -= 1
            continue
        if not stripped:
            i -= 1
            continue
        break
    attrs.reverse()
    return "\n".join(attrs)


def _iter_methods(test_path: Path) -> list[MethodBlock]:
    lines = test_path.read_text(encoding="utf-8").splitlines()
    methods: list[MethodBlock] = []
    i = 0
    while i < len(lines):
        match = METHOD_RE.match(lines[i])
        if match is None:
            i += 1
            continue
        name = match.group(1)
        end = _find_method_end(lines, i)
        methods.append(
            MethodBlock(
                name=name,
                start_line=i + 1,
                attributes=_collect_attributes(lines, i),
                body="\n".join(lines[i : end + 1]),
                source=str(test_path),
            )
        )
        i = end + 1
    return methods


def _normalize_versions(versions: list[str]) -> tuple[str, ...]:
    ordered: list[str] = []
    seen: set[str] = set()
    for version in versions:
        if version not in VALID_VERSIONS:
            continue
        if version in seen:
            continue
        seen.add(version)
        ordered.append(version)
    return tuple(ordered)


def _resolve_symbol(
    symbol: str,
    symbol_to_resource: dict[str, str],
    dotted_exact: dict[str, str],
    dotted_lower: dict[str, str],
) -> str | None:
    candidates = [symbol]
    if symbol.startswith("TestAssets."):
        candidates.append(symbol.removeprefix("TestAssets."))
    if "." in symbol:
        candidates.append(symbol.split(".")[-1])

    for candidate in candidates:
        resource = symbol_to_resource.get(candidate)
        if resource is None:
            continue
        relpath = _resolve_resource_to_relpath(resource, dotted_exact, dotted_lower)
        if relpath is not None:
            return relpath
    return None


def _parse_validator_aliases(test_path: Path) -> dict[str, str]:
    text = test_path.read_text(encoding="utf-8")
    aliases: dict[str, str] = {}

    for match in VALIDATOR_ALIAS_VERSION_RE.finditer(text):
        alias = match.group(1)
        version = match.group(2)
        if version in VALID_VERSIONS:
            aliases[alias] = version

    for match in VALIDATOR_ALIAS_DEFAULT_RE.finditer(text):
        alias = match.group(1)
        aliases.setdefault(alias, "Office2007")

    return aliases


def _infer_version_from_validator_name(name: str) -> str | None:
    if name in VALID_VERSIONS:
        return name
    match = re.search(r"^O(\d{2})", name)
    if match is None:
        return None
    return VALIDATOR_NAME_VERSION_MAP.get(match.group(1))


def _build_expectation(
    relpath: str,
    kind: str,
    versions: tuple[str, ...],
    scenario: str,
    source: MethodBlock,
    expected_count: int | None = None,
    expected_counts: tuple[int, ...] = (),
) -> ExtractedExpectation:
    return ExtractedExpectation(
        relpath=relpath,
        kind=kind,
        validator_versions=versions,
        scenario=scenario,
        expected_error_count=expected_count,
        expected_error_counts=expected_counts,
        source_file=Path(source.source).name,
        source_method=source.name,
        source_line=source.start_line,
    )


def _classify_method_scenario(method: MethodBlock) -> str:
    """Classify whether expectations come from base-file validation or a mutated scenario."""
    if MAX_ERRORS_SIGNAL_RE.search(method.body):
        return "mutation"

    if MUTATION_NAME_RE.search(method.name):
        return "mutation"

    validate_match = VALIDATE_CALL_RE.search(method.body)
    prefix = method.body if validate_match is None else method.body[: validate_match.start()]
    if MUTATION_SIGNAL_RE.search(prefix):
        return "mutation"
    return "base"


def _collect_package_variables(method_body: str) -> set[str]:
    return {match.group(1) for match in PACKAGE_OPEN_RE.finditer(method_body)}


def _extract_first_argument(call_source: str, start_index: int) -> tuple[str, int]:
    depth = 0
    idx = start_index
    while idx < len(call_source):
        ch = call_source[idx]
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            if depth == 0:
                break
            depth -= 1
        elif ch == "," and depth == 0:
            break
        idx += 1
    return call_source[start_index:idx].strip(), idx


def _normalize_validate_target(expression: str) -> str:
    compact = "".join(expression.split())
    while compact.startswith("("):
        close = compact.find(")")
        if close <= 0:
            break
        compact = compact[close + 1 :]
    return compact.lstrip("!")


def _is_package_validate_target(expression: str, package_variables: set[str]) -> bool:
    target = _normalize_validate_target(expression)
    if not target:
        return False
    if "." in target or "[" in target or "(" in target:
        return False
    return target in package_variables


def _extract_from_method(
    method: MethodBlock,
    symbol_to_resource: dict[str, str],
    dotted_exact: dict[str, str],
    dotted_lower: dict[str, str],
    validator_aliases: dict[str, str],
) -> list[ExtractedExpectation]:
    results: list[ExtractedExpectation] = []
    scenario = _classify_method_scenario(method)
    package_variables = _collect_package_variables(method.body)

    file_symbols = sorted(set(GET_STREAM_RE.findall(method.body)))
    resolved_files = [
        _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        for symbol in file_symbols
    ]
    resolved_files = [path for path in resolved_files if path is not None]
    single_file = resolved_files[0] if len(resolved_files) == 1 else None
    inline_files = []
    for match in INLINE_FIRST_ARG_RE.finditer(method.attributes):
        symbol = match.group(1)
        if symbol.startswith("FileFormatVersions."):
            continue
        relpath = _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        if relpath is not None:
            inline_files.append(relpath)
    inline_files = list(dict.fromkeys(inline_files))

    versions = list(VALIDATOR_VERSION_RE.findall(method.body))
    if VALIDATOR_DEFAULT_RE.search(method.body):
        versions.append("Office2007")
    normalized_versions = _normalize_versions(versions)
    has_empty_validate = "Assert.Empty(" in method.body and "Validate(" in method.body
    seen_rows: set[tuple[str, str, tuple[str, ...], str, int | None, tuple[int, ...]]] = set()

    def add_expectation(
        relpath: str,
        *,
        kind: str,
        versions_payload: tuple[str, ...],
        expected_count: int | None = None,
        expected_counts: tuple[int, ...] = (),
    ) -> None:
        key = (relpath, kind, versions_payload, scenario, expected_count, expected_counts)
        if key in seen_rows:
            return
        seen_rows.add(key)
        results.append(
            _build_expectation(
                relpath=relpath,
                kind=kind,
                versions=versions_payload,
                scenario=scenario,
                expected_count=expected_count,
                expected_counts=expected_counts,
                source=method,
            )
        )

    def resolve_validator_versions(validator_name: str) -> tuple[str, ...]:
        if validator_name in validator_aliases:
            return (validator_aliases[validator_name],)
        inferred = _infer_version_from_validator_name(validator_name)
        if inferred is not None:
            return (inferred,)
        if len(normalized_versions) == 1:
            return normalized_versions
        return ()

    for match in INLINE_FILE_COUNT_RE.finditer(method.attributes):
        symbol = match.group(1)
        expected = int(match.group(2))
        relpath = _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        if relpath is None:
            continue
        add_expectation(
            relpath,
            kind="inline_file_count",
            versions_payload=normalized_versions,
            expected_count=expected,
        )

    for match in INLINE_FILE_ONLY_RE.finditer(method.attributes):
        symbol = match.group(1)
        if symbol.startswith("FileFormatVersions."):
            continue
        relpath = _resolve_symbol(symbol, symbol_to_resource, dotted_exact, dotted_lower)
        if relpath is None:
            continue
        if has_empty_validate and len(normalized_versions) == 1:
            add_expectation(
                relpath,
                kind="inline_file_assert_empty",
                versions_payload=normalized_versions,
                expected_count=0,
            )

    for match in INLINE_VERSION_COUNT_RE.finditer(method.attributes):
        version = match.group(1)
        expected = int(match.group(2))
        if single_file is None:
            continue
        add_expectation(
            single_file,
            kind="inline_version_count",
            versions_payload=(version,),
            expected_count=expected,
        )

    for match in INLINE_VERSION_ONLY_RE.finditer(method.attributes):
        if single_file is None or not has_empty_validate:
            continue
        version = match.group(1)
        add_expectation(
            single_file,
            kind="inline_version_assert_empty",
            versions_payload=(version,),
            expected_count=0,
        )

    if (
        has_empty_validate
        and single_file is not None
        and len(normalized_versions) == 1
    ):
        add_expectation(
            single_file,
            kind="assert_empty_single_version",
            versions_payload=normalized_versions,
            expected_count=0,
        )

    allowed_match = ASSERT_TRUE_ALLOWED_COUNTS_RE.search(method.body)
    if (
        allowed_match is not None
        and single_file is not None
        and len(normalized_versions) >= 1
    ):
        a = int(allowed_match.group(1))
        b = int(allowed_match.group(2))
        add_expectation(
            single_file,
            kind="assert_true_allowed_counts",
            versions_payload=normalized_versions,
            expected_counts=tuple(sorted({a, b})),
        )

    helper_match = XLSX_HELPER_CALL_RE.search(method.body)
    if helper_match is not None and single_file is not None:
        a = int(helper_match.group(1))
        b = int(helper_match.group(2))
        add_expectation(
            single_file,
            kind="helper_allowed_counts",
            versions_payload=("Office2007", "Office2010", "Office2013"),
            expected_counts=tuple(sorted({a, b})),
        )

    target_files = [single_file] if single_file is not None else inline_files
    if target_files:
        package_invocation_versions: list[str] = []
        assignments: dict[str, list[tuple[int, str, bool]]] = defaultdict(list)
        for match in ASSIGN_VALIDATE_RE.finditer(method.body):
            var_name = match.group(1)
            validator_name = match.group(2)
            arg_expression, _ = _extract_first_argument(method.body, match.end())
            is_package_target = _is_package_validate_target(arg_expression, package_variables)
            assignments[var_name].append((match.start(), validator_name, is_package_target))
            if is_package_target:
                package_invocation_versions.extend(resolve_validator_versions(validator_name))

        def latest_assignment(var_name: str, before_pos: int) -> tuple[str, bool] | None:
            entries = assignments.get(var_name)
            if not entries:
                return None
            latest: tuple[str, bool] | None = None
            for pos, validator_name, is_package_target in entries:
                if pos >= before_pos:
                    break
                latest = (validator_name, is_package_target)
            return latest

        for match in ASSERT_EMPTY_DIRECT_RE.finditer(method.body):
            arg_expression, _ = _extract_first_argument(method.body, match.end())
            if not _is_package_validate_target(arg_expression, package_variables):
                continue
            versions_payload = resolve_validator_versions(match.group(1))
            if not versions_payload:
                continue
            package_invocation_versions.extend(versions_payload)
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_empty_validator",
                    versions_payload=versions_payload,
                    expected_count=0,
                )

        for match in ASSERT_SINGLE_DIRECT_RE.finditer(method.body):
            arg_expression, _ = _extract_first_argument(method.body, match.end())
            if not _is_package_validate_target(arg_expression, package_variables):
                continue
            versions_payload = resolve_validator_versions(match.group(1))
            if not versions_payload:
                continue
            package_invocation_versions.extend(versions_payload)
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_single_validator",
                    versions_payload=versions_payload,
                    expected_count=1,
                )

        for match in ASSERT_EQUAL_DIRECT_COUNT_RE.finditer(method.body):
            arg_expression, arg_end = _extract_first_argument(method.body, match.end())
            if not _is_package_validate_target(arg_expression, package_variables):
                continue
            if re.match(r"\s*\)\s*\.Count\s*\(", method.body[arg_end:]) is None:
                continue
            versions_payload = resolve_validator_versions(match.group(2))
            if not versions_payload:
                continue
            package_invocation_versions.extend(versions_payload)
            expected = int(match.group(1))
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_equal_validator_count",
                    versions_payload=versions_payload,
                    expected_count=expected,
                )

        for match in ASSERT_NOTNULL_DIRECT_RE.finditer(method.body):
            arg_expression, _ = _extract_first_argument(method.body, match.end())
            if not _is_package_validate_target(arg_expression, package_variables):
                continue
            versions_payload = resolve_validator_versions(match.group(1))
            if not versions_payload:
                continue
            package_invocation_versions.extend(versions_payload)
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_not_null_validator",
                    versions_payload=versions_payload,
                )

        for match in ASSERT_EMPTY_VAR_RE.finditer(method.body):
            var_name = match.group(1)
            assignment = latest_assignment(var_name, match.start())
            if assignment is None:
                continue
            validator_name, is_package_target = assignment
            if not is_package_target:
                continue
            versions_payload = resolve_validator_versions(validator_name)
            if not versions_payload:
                continue
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_empty_validator",
                    versions_payload=versions_payload,
                    expected_count=0,
                )

        for match in ASSERT_SINGLE_VAR_RE.finditer(method.body):
            var_name = match.group(1)
            assignment = latest_assignment(var_name, match.start())
            if assignment is None:
                continue
            validator_name, is_package_target = assignment
            if not is_package_target:
                continue
            versions_payload = resolve_validator_versions(validator_name)
            if not versions_payload:
                continue
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_single_validator",
                    versions_payload=versions_payload,
                    expected_count=1,
                )

        for match in ASSERT_EQUAL_VAR_COUNT_RE.finditer(method.body):
            var_name = match.group(2)
            assignment = latest_assignment(var_name, match.start())
            if assignment is None:
                continue
            validator_name, is_package_target = assignment
            if not is_package_target:
                continue
            versions_payload = resolve_validator_versions(validator_name)
            if not versions_payload:
                continue
            expected = int(match.group(1))
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_equal_validator_count",
                    versions_payload=versions_payload,
                    expected_count=expected,
                )

        for match in ASSERT_NOTNULL_VAR_RE.finditer(method.body):
            var_name = match.group(1)
            assignment = latest_assignment(var_name, match.start())
            if assignment is None:
                continue
            validator_name, is_package_target = assignment
            if not is_package_target:
                continue
            versions_payload = resolve_validator_versions(validator_name)
            if not versions_payload:
                continue
            for relpath in target_files:
                add_expectation(
                    relpath,
                    kind="assert_not_null_validator",
                    versions_payload=versions_payload,
                )

        if package_invocation_versions:
            deduped_versions = _normalize_versions(package_invocation_versions)
            for version in deduped_versions:
                for relpath in target_files:
                    add_expectation(
                        relpath,
                        kind="validate_invocation",
                        versions_payload=(version,),
                    )

    return results


def _aggregate_expectations(items: list[ExtractedExpectation]) -> dict[str, list[dict[str, Any]]]:
    by_file: dict[str, dict[str, dict[str, Any]]] = defaultdict(dict)
    for item in items:
        key_payload = {
            "kind": item.kind,
            "validator_versions": list(item.validator_versions),
            "scenario": item.scenario,
        }
        if item.expected_error_count is not None:
            key_payload["expected_error_count"] = item.expected_error_count
        if item.expected_error_counts:
            key_payload["expected_error_counts"] = list(item.expected_error_counts)
        key = json.dumps(key_payload, sort_keys=True)

        existing = by_file[item.relpath].get(key)
        if existing is None:
            by_file[item.relpath][key] = {
                **key_payload,
                "source_tests": [
                    {
                        "file": item.source_file,
                        "name": item.source_method,
                        "line": item.source_line,
                    }
                ],
            }
            continue

        source_list = existing["source_tests"]
        assert isinstance(source_list, list)
        source_list.append(
            {
                "file": item.source_file,
                "name": item.source_method,
                "line": item.source_line,
            }
        )

    finalized: dict[str, list[dict[str, Any]]] = {}
    for relpath, grouped in by_file.items():
        values = list(grouped.values())
        values.sort(
            key=lambda row: (
                str(row.get("scenario", "base")),
                str(row.get("kind", "")),
                ",".join(row.get("validator_versions", [])),
                int(row.get("expected_error_count", -1)),
            )
        )
        for value in values:
            tests = value["source_tests"]
            assert isinstance(tests, list)
            tests.sort(key=lambda row: (str(row.get("file", "")), int(row.get("line", 0))))
        finalized[relpath] = values
    return finalized


def _collect_expectations(
    sdk_root: Path,
    symbol_to_resource: dict[str, str],
    dotted_exact: dict[str, str],
    dotted_lower: dict[str, str],
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, int]]:
    all_items: list[ExtractedExpectation] = []
    stats = {
        "methods_scanned": 0,
        "raw_expectations": 0,
        "raw_expectations_base": 0,
        "raw_expectations_mutation": 0,
        "resolved_expectations": 0,
        "resolved_expectations_base": 0,
        "resolved_expectations_mutation": 0,
        "files_with_expectations": 0,
    }
    tests_root = sdk_root / TESTS_DIR
    for test_file in sorted(tests_root.rglob("*.cs")):
        methods = _iter_methods(test_file)
        validator_aliases = _parse_validator_aliases(test_file)
        stats["methods_scanned"] += len(methods)
        for method in methods:
            extracted = _extract_from_method(
                method,
                symbol_to_resource,
                dotted_exact,
                dotted_lower,
                validator_aliases,
            )
            stats["raw_expectations"] += len(extracted)
            stats["raw_expectations_base"] += sum(
                1 for item in extracted if item.scenario == "base"
            )
            stats["raw_expectations_mutation"] += sum(
                1 for item in extracted if item.scenario == "mutation"
            )
            all_items.extend(extracted)

    aggregated = _aggregate_expectations(all_items)
    stats["resolved_expectations"] = sum(len(values) for values in aggregated.values())
    stats["resolved_expectations_base"] = sum(
        1
        for values in aggregated.values()
        for row in values
        if row.get("scenario", "base") == "base"
    )
    stats["resolved_expectations_mutation"] = sum(
        1
        for values in aggregated.values()
        for row in values
        if row.get("scenario", "base") == "mutation"
    )
    stats["files_with_expectations"] = len(aggregated)
    return aggregated, stats


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract SDK expectations into corpus manifest.")
    parser.add_argument(
        "--sdk-root",
        type=Path,
        default=DEFAULT_SDK_ROOT,
        help=f"Path to Open XML SDK clone (default: {DEFAULT_SDK_ROOT})",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        default=DEFAULT_MANIFEST,
        help=f"Input manifest (default: {DEFAULT_MANIFEST})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output manifest path (default: overwrite --manifest)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Analyze and print stats without writing output.",
    )
    args = parser.parse_args()

    sdk_root = args.sdk_root.resolve()
    manifest_path = args.manifest.resolve()
    output_path = args.output.resolve() if args.output is not None else manifest_path

    if not manifest_path.exists():
        print(f"Manifest not found: {manifest_path}")
        return 2
    if not (sdk_root / ASSETS_DIR).exists():
        print(f"SDK assets directory not found: {sdk_root / ASSETS_DIR}")
        return 2
    if not (sdk_root / TESTS_DIR).exists():
        print(f"SDK tests directory not found: {sdk_root / TESTS_DIR}")
        return 2

    manifest = _load_manifest(manifest_path)
    dotted_exact, dotted_lower = _build_relpath_indexes(manifest)
    symbol_to_resource = _parse_constant_symbols(sdk_root / ASSETS_DIR)
    expectations_by_file, stats = _collect_expectations(
        sdk_root=sdk_root,
        symbol_to_resource=symbol_to_resource,
        dotted_exact=dotted_exact,
        dotted_lower=dotted_lower,
    )

    files = manifest.get("files", [])
    if not isinstance(files, list):
        print("Manifest has invalid 'files' payload.")
        return 2

    for entry in files:
        relpath = entry.get("source_relpath")
        if not isinstance(relpath, str):
            continue
        entry["expectations"] = expectations_by_file.get(relpath, [])

    manifest["expectation_stats"] = stats
    manifest["expectation_extraction"] = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "sdk_root": str(sdk_root),
        "script": "scripts/corpus/extract_sdk_expectations.py",
    }
    manifest["matched_expectation_files"] = stats["files_with_expectations"]

    print(f"SDK root: {sdk_root}")
    print(f"Manifest: {manifest_path}")
    print(f"Methods scanned: {stats['methods_scanned']}")
    print(f"Raw expectations: {stats['raw_expectations']}")
    print(f"Raw expectations (base): {stats['raw_expectations_base']}")
    print(f"Raw expectations (mutation): {stats['raw_expectations_mutation']}")
    print(f"Resolved expectations: {stats['resolved_expectations']}")
    print(f"Resolved expectations (base): {stats['resolved_expectations_base']}")
    print(f"Resolved expectations (mutation): {stats['resolved_expectations_mutation']}")
    print(f"Files with expectations: {stats['files_with_expectations']}")

    if args.dry_run:
        print("Dry run enabled, nothing written.")
        return 0

    output_path.parent.mkdir(parents=True, exist_ok=True)
    _save_manifest(output_path, manifest)
    print(f"Wrote manifest: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
