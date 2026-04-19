"""Runtime interpretation of Open XML SDK data files.

This module provides direct access to the SDK schema and schematron definitions
without code generation. The JSON files are the single source of truth.
"""

from __future__ import annotations

from openxml_audit.codegen.constraint_bridge import (
    convert_element_type,
    get_element_constraint,
    get_sdk_element_info,
)
from openxml_audit.codegen.schema_loader import (
    SchemaRegistry,
    SdkAttribute,
    SdkElementType,
    SdkParticle,
    SdkSchema,
    get_xsd_type_name,
)
from openxml_audit.codegen.schema_loader import (
    get_registry as get_schema_registry,
)
from openxml_audit.codegen.schematron_loader import (
    ParsedSchematron,
    SchematronRegistry,
    SchematronType,
)
from openxml_audit.codegen.schematron_loader import (
    get_registry as get_schematron_registry,
)

__all__ = [
    "SchemaRegistry",
    "SdkAttribute",
    "SdkElementType",
    "SdkParticle",
    "SdkSchema",
    "get_schema_registry",
    "get_xsd_type_name",
    "ParsedSchematron",
    "SchematronRegistry",
    "SchematronType",
    "get_schematron_registry",
    "get_element_constraint",
    "get_sdk_element_info",
    "convert_element_type",
]
