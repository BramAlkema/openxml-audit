"""Load and interpret Open XML SDK schema JSON files at runtime.

This module provides direct access to the SDK schema definitions without
code generation. The JSON files are the single source of truth.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterator

if TYPE_CHECKING:
    from openxml_audit.schema.types import XsdTypeValidator

# Path to SDK data files
DATA_DIR = Path(__file__).parent.parent.parent.parent / "data" / "openxml"
SCHEMAS_DIR = DATA_DIR / "schemas"


@dataclass
class SdkAttribute:
    """Attribute definition from SDK schema."""

    qname: str  # e.g., ":id" (no prefix = no namespace) or "r:id"
    property_name: str
    type_name: str  # e.g., "UInt32Value", "BooleanValue", "StringValue"
    required: bool = False
    validators: list[dict[str, Any]] = field(default_factory=list)

    @property
    def local_name(self) -> str:
        """Get the local name (after the colon)."""
        if ":" in self.qname:
            return self.qname.split(":", 1)[1]
        return self.qname

    @property
    def prefix(self) -> str | None:
        """Get the namespace prefix (before the colon), or None if no prefix."""
        if ":" in self.qname:
            prefix = self.qname.split(":", 1)[0]
            return prefix if prefix else None
        return None


@dataclass
class SdkParticle:
    """Particle (content model) definition from SDK schema."""

    kind: str  # "Sequence", "Choice", "Group", "Any", or element name reference
    items: list[SdkParticle] = field(default_factory=list)
    min_occurs: int = 1
    max_occurs: int = 1  # -1 means unbounded
    name: str | None = None  # For element references like "a:CT_OfficeArtExtensionList/a:extLst"

    @classmethod
    def from_json(cls, data: dict[str, Any]) -> SdkParticle:
        """Parse a particle from SDK JSON."""
        kind = data.get("Kind", "Element")

        # Parse occurrence constraints
        # Default: if Occurs is absent or only has Max, element is optional (minOccurs=0)
        # Only if Min is explicitly set do we use that value
        min_occurs = 0  # Default to optional
        max_occurs = 1
        if "Occurs" in data:
            has_empty = False
            has_max = False
            for occur in data["Occurs"]:
                if not occur:
                    has_empty = True
                    continue
                if "Min" in occur:
                    min_occurs = occur["Min"]
                if "Max" in occur:
                    has_max = True
                    max_val = occur["Max"]
                    max_occurs = -1 if max_val == 0 else max_val  # 0 means unbounded in SDK
            if has_empty:
                # Empty occurrence means unbounded in SDK JSON.
                max_occurs = -1
            if not has_max:
                # Min-only occurrences default to unbounded in SDK JSON.
                max_occurs = -1
        else:
            # No Occurs specified - element is required once (min=1, max=1)
            min_occurs = 1

        # Parse child items
        items = []
        if "Items" in data:
            items = [cls.from_json(item) for item in data["Items"]]

        return cls(
            kind=kind,
            items=items,
            min_occurs=min_occurs,
            max_occurs=max_occurs,
            name=data.get("Name"),
        )


@dataclass
class SdkElementType:
    """Element type definition from SDK schema."""

    name: str  # e.g., "p:CT_Slide/p:sld"
    class_name: str
    base_class: str | None = None
    is_abstract: bool = False
    is_derived: bool = False
    is_leaf_element: bool = False
    summary: str | None = None
    attributes: list[SdkAttribute] = field(default_factory=list)
    particle: SdkParticle | None = None
    children: list[dict[str, Any]] = field(default_factory=list)

    @property
    def type_name(self) -> str:
        """Get the complex type name (e.g., 'CT_Slide')."""
        if "/" in self.name:
            return self.name.split("/")[0].split(":")[-1]
        return self.name.split(":")[-1]

    @property
    def element_name(self) -> str | None:
        """Get the element local name (e.g., 'sld'), or None for abstract types."""
        if "/" in self.name:
            elem_part = self.name.split("/")[1]
            if elem_part:
                return elem_part.split(":")[-1]
        return None

    @property
    def element_prefix(self) -> str | None:
        """Get the element namespace prefix (e.g., 'p')."""
        if "/" in self.name:
            elem_part = self.name.split("/")[1]
            if elem_part and ":" in elem_part:
                return elem_part.split(":")[0]
        return None

    @classmethod
    def from_json(cls, data: dict[str, Any]) -> SdkElementType:
        """Parse an element type from SDK JSON."""
        # Parse attributes
        attributes = []
        for attr_data in data.get("Attributes", []):
            required = any(
                v.get("Name") == "RequiredValidator"
                for v in attr_data.get("Validators", [])
            )
            qname = attr_data["QName"]
            # PropertyName may be missing - derive from QName
            property_name = attr_data.get("PropertyName")
            if not property_name:
                # Use local part of QName as property name
                property_name = qname.split(":")[-1] if ":" in qname else qname.lstrip(":")
            attributes.append(SdkAttribute(
                qname=qname,
                property_name=property_name,
                type_name=attr_data.get("Type", "StringValue"),
                required=required,
                validators=attr_data.get("Validators", []),
            ))

        # Parse particle (content model)
        particle = None
        if "Particle" in data:
            particle = SdkParticle.from_json(data["Particle"])

        return cls(
            name=data["Name"],
            class_name=data["ClassName"],
            base_class=data.get("BaseClass"),
            is_abstract=data.get("IsAbstract", False),
            is_derived=data.get("IsDerived", False),
            is_leaf_element=data.get("IsLeafElement", False),
            summary=data.get("Summary"),
            attributes=attributes,
            particle=particle,
            children=data.get("Children", []),
        )


@dataclass
class SdkSchema:
    """A schema definition from SDK JSON."""

    target_namespace: str
    types: list[SdkElementType]
    _types_by_name: dict[str, SdkElementType] = field(default_factory=dict, repr=False)

    def __post_init__(self) -> None:
        """Build lookup indexes."""
        for t in self.types:
            self._types_by_name[t.name] = t

    def get_type(self, name: str) -> SdkElementType | None:
        """Get a type by its full name."""
        return self._types_by_name.get(name)

    def get_element_types(self) -> Iterator[SdkElementType]:
        """Get all concrete (non-abstract) element types."""
        for t in self.types:
            if not t.is_abstract and t.element_name:
                yield t

    @classmethod
    def from_json(cls, data: dict[str, Any]) -> SdkSchema:
        """Parse a schema from SDK JSON."""
        types = [SdkElementType.from_json(t) for t in data.get("Types", [])]
        return cls(
            target_namespace=data["TargetNamespace"],
            types=types,
        )


class SchemaRegistry:
    """Registry of all SDK schemas with lookup capabilities."""

    def __init__(self) -> None:
        self._schemas: dict[str, SdkSchema] = {}  # namespace -> schema
        self._types: dict[str, SdkElementType] = {}  # full name -> type
        self._elements: dict[str, SdkElementType] = {}  # {ns}localname -> type
        self._elements_by_tag: dict[str, list[SdkElementType]] = {}
        self._prefixes: dict[str, str] = {}  # prefix -> namespace URI
        self._loaded = False

    def load(self) -> None:
        """Load all schemas from disk."""
        if self._loaded:
            return

        # Load namespace mappings
        ns_file = DATA_DIR / "namespaces.json"
        if ns_file.exists():
            with open(ns_file) as f:
                for ns in json.load(f):
                    if ns.get("Prefix") and ns.get("Uri"):
                        self._prefixes[ns["Prefix"]] = ns["Uri"]

        # Load all schema files
        if SCHEMAS_DIR.exists():
            for schema_file in SCHEMAS_DIR.glob("*.json"):
                try:
                    with open(schema_file) as f:
                        schema = SdkSchema.from_json(json.load(f))
                        self._schemas[schema.target_namespace] = schema

                        # Index types
                        for t in schema.types:
                            self._types[t.name] = t

                            # Index concrete elements by qualified name
                            if not t.is_abstract and t.element_name:
                                prefix = t.element_prefix
                                ns = self._prefixes.get(prefix, "") if prefix else schema.target_namespace
                                qname = f"{{{ns}}}{t.element_name}"
                                self._elements_by_tag.setdefault(qname, []).append(t)

                                # Prefer "primary" element types:
                                # - CompositeElement over LeafElement
                                # - Types with particle (content model)
                                # - Types with more attributes
                                existing = self._elements.get(qname)
                                if existing:
                                    # Score based on "richness"
                                    def particle_size(particle: SdkParticle | None) -> int:
                                        if particle is None:
                                            return 0
                                        total = len(particle.items)
                                        for item in particle.items:
                                            total += particle_size(item)
                                        return total

                                    def score(elem: SdkElementType) -> int:
                                        s = 0
                                        if not elem.is_leaf_element:
                                            s += 100  # Composite elements are primary
                                        if elem.particle:
                                            s += 50  # Has content model
                                        s += particle_size(elem.particle)
                                        s += len(elem.attributes)
                                        return s

                                    if score(t) <= score(existing):
                                        continue  # Keep existing

                                self._elements[qname] = t

                except (json.JSONDecodeError, KeyError) as e:
                    # Log but continue loading other schemas
                    print(f"Warning: Failed to load {schema_file}: {e}")

        self._loaded = True

    def get_namespace(self, prefix: str) -> str | None:
        """Get namespace URI for a prefix."""
        self.load()
        return self._prefixes.get(prefix)

    def get_schema(self, namespace: str) -> SdkSchema | None:
        """Get schema for a namespace."""
        self.load()
        return self._schemas.get(namespace)

    def get_type(self, name: str) -> SdkElementType | None:
        """Get a type by its full SDK name (e.g., 'p:CT_Slide/p:sld')."""
        self.load()
        return self._types.get(name)

    def get_element_type(self, namespace: str, local_name: str) -> SdkElementType | None:
        """Get element type by namespace and local name."""
        self.load()
        return self._elements.get(f"{{{namespace}}}{local_name}")

    def get_element_type_by_tag(self, tag: str) -> SdkElementType | None:
        """Get element type by Clark notation tag (e.g., '{namespace}localname')."""
        self.load()
        return self._elements.get(tag)

    def get_element_type_candidates(self, tag: str) -> list[SdkElementType]:
        """Get all element type candidates for a tag."""
        self.load()
        return list(self._elements_by_tag.get(tag, []))

    def list_schemas(self) -> list[str]:
        """List all loaded schema namespaces."""
        self.load()
        return list(self._schemas.keys())

    def count_elements(self) -> int:
        """Count total number of concrete elements."""
        self.load()
        return len(self._elements)

    def count_types(self) -> int:
        """Count total number of types."""
        self.load()
        return len(self._types)


# Global registry instance
_registry: SchemaRegistry | None = None


def get_registry() -> SchemaRegistry:
    """Get the global schema registry."""
    global _registry
    if _registry is None:
        _registry = SchemaRegistry()
    return _registry


# Type mapping from SDK types to XSD built-in types
SDK_TYPE_MAP: dict[str, str] = {
    "StringValue": "string",
    "BooleanValue": "boolean",
    "Int16Value": "short",
    "Int32Value": "int",
    "Int64Value": "long",
    "UInt16Value": "unsignedShort",
    "UInt32Value": "unsignedInt",
    "UInt64Value": "unsignedLong",
    "ByteValue": "byte",
    "SByteValue": "unsignedByte",
    "SingleValue": "float",
    "DoubleValue": "double",
    "DecimalValue": "decimal",
    "DateTimeValue": "dateTime",
    "HexBinaryValue": "hexBinary",
    "Base64BinaryValue": "base64Binary",
    # Enum types map to string with enumeration
    "EnumValue": "string",
}


def get_xsd_type_name(sdk_type: str) -> str:
    """Map SDK type name to XSD built-in type name."""
    # Handle EnumValue<T> pattern
    if sdk_type.startswith("EnumValue<"):
        return "string"
    return SDK_TYPE_MAP.get(sdk_type, "string")
