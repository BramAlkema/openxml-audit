"""Parity normalization helpers for SDK-style comparison."""

from __future__ import annotations

import re

from openxml_audit.errors import ValidationError, ValidationErrorType


def normalize_description(text: str) -> str:
    """Normalize volatile values in error descriptions."""
    value = re.sub(r"'[^']*'", "'<value>'", text)
    value = re.sub(r"\b\d+(?:\.\d+)?\b", "<n>", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalize_part_uri(part_uri: str) -> str:
    """Normalize part URI for stable tuple comparison."""
    if not part_uri:
        return "/"
    return part_uri if part_uri.startswith("/") else f"/{part_uri}"


def normalize_path(path: str) -> str:
    """Normalize xpath-like paths to a stable SDK-style shape."""
    if not path:
        return "/"
    segments = [segment for segment in path.split("/") if segment]
    normalized: list[str] = []
    for segment in segments:
        if re.search(r"\[\d+\]$", segment):
            normalized.append(segment)
        else:
            normalized.append(f"{segment}[1]")
    return "/" + "/".join(normalized)


def normalize_error_id(error: ValidationError) -> str:
    """Map internal errors to stable SDK-style ID families."""
    if error.id:
        return error.id

    description = error.description.lower()
    error_type = error.error_type

    if error_type == ValidationErrorType.SCHEMA:
        if "required choice element is missing" in description:
            return "Sch_IncompleteContentExpectingComplex"
        if "required element" in description and "is missing" in description:
            return "Sch_IncompleteContentExpectingComplex"
        if "unexpected element" in description:
            return "Sch_UnexpectedElementContentExpectingComplex"
        if "required attribute" in description and "is missing" in description:
            return "Sch_MissRequiredAttribute"
        if "invalid value for attribute" in description:
            return "Sch_AttributeValueDataTypeDetailed"
        if "attribute" in description and "is not declared" in description:
            return "Sch_UndeclaredAttribute"
        return "Sch_SchemaError"

    if error_type == ValidationErrorType.SEMANTIC:
        if "duplicate id" in description:
            return "Sem_UniqueId"
        if "referenced by" in description and "does not exist" in description:
            return "Sem_MissingRelationshipReference"
        if "missing required relationship type" in description:
            return "Sem_MissingRelationshipType"
        return "Sem_SemanticError"

    if error_type == ValidationErrorType.RELATIONSHIP:
        return "Pkg_RelationshipError"
    if error_type == ValidationErrorType.PACKAGE:
        return "Pkg_PackageError"
    if error_type == ValidationErrorType.BINARY:
        return "Pkg_BinaryPayloadError"
    if error_type == ValidationErrorType.MARKUP_COMPATIBILITY:
        return "Mc_MarkupCompatibilityError"

    return "Unknown"


def normalize_error_tuple(error: ValidationError) -> dict[str, str]:
    """Build normalized tuple fields used by parity comparison."""
    normalized_id = normalize_error_id(error)
    normalized_error_type = error.error_type.value
    normalized_part = normalize_part_uri(error.part_uri)
    normalized_path = normalize_path(error.path)
    normalized_description = normalize_description(error.description)
    family_key = "|".join(
        (
            normalized_id,
            normalized_error_type,
            normalized_part,
            normalized_path,
            normalized_description,
        )
    )
    return {
        "id": normalized_id,
        "error_type": normalized_error_type,
        "part": normalized_part,
        "path": normalized_path,
        "description": normalized_description,
        "family_key": family_key,
    }
