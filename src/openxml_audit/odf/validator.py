"""ODF validator skeleton."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from lxml import etree

from openxml_audit.errors import (
    FileFormat,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.odf.package import OdfPackage


class OdfValidator:
    """Validate ODF packages using manifest + schema rules."""

    OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    CORE_ROOTS = {
        "content.xml": "document-content",
        "styles.xml": "document-styles",
        "meta.xml": "document-meta",
        "settings.xml": "document-settings",
    }
    CONTENT_BODY_BY_MIMETYPE = (
        ("application/vnd.oasis.opendocument.text", "text"),
        ("application/vnd.oasis.opendocument.spreadsheet", "spreadsheet"),
        ("application/vnd.oasis.opendocument.presentation", "presentation"),
    )

    def __init__(
        self,
        file_format: FileFormat = FileFormat.ODF_1_3,
        strict: bool = True,
        *,
        relaxng_validation: bool = False,
        relaxng_schemas: Mapping[str, str | Path] | None = None,
    ):
        if relaxng_validation and not relaxng_schemas:
            raise ValueError(
                "relaxng_validation requires relaxng_schemas mapping "
                "(for example {'content.xml': '/path/to/content.rng'})"
            )

        self._file_format = file_format
        self._strict = strict
        self._relaxng_validation = relaxng_validation
        self._relaxng_schemas = dict(relaxng_schemas or {})
        self._relaxng_cache: dict[Path, etree.RelaxNG] = {}

    @staticmethod
    def _normalize_part_uri(part_path: str) -> str:
        return part_path if part_path.startswith("/") else f"/{part_path}"

    @staticmethod
    def _normalize_part_key(part_path: str) -> str:
        return part_path[1:] if part_path.startswith("/") else part_path

    def _collect_xml_parts(self, package: OdfPackage) -> list[str]:
        parts: set[str] = set(package.list_xml_parts())
        for core in ("content.xml", "styles.xml", "meta.xml", "settings.xml"):
            if package.has_part(core):
                parts.add(core)
        return sorted(parts)

    def _parse_xml_parts(
        self,
        package: OdfPackage,
    ) -> tuple[dict[str, etree._Element], list[ValidationError]]:
        parsed_parts: dict[str, etree._Element] = {}
        errors: list[ValidationError] = []
        for part in self._collect_xml_parts(package):
            content = package.get_part_content(part)
            if content is None:
                continue
            try:
                parsed_parts[part] = etree.fromstring(content)
            except etree.XMLSyntaxError as exc:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.SCHEMA,
                        description=f"XML parse error: {exc}",
                        part_uri=self._normalize_part_uri(part),
                        severity=ValidationSeverity.ERROR,
                    )
                )
        return parsed_parts, errors

    def _schema_path_for_part(self, part: str) -> Path | None:
        normalized = self._normalize_part_key(part)
        schema = self._relaxng_schemas.get(part)
        if schema is None:
            schema = self._relaxng_schemas.get(normalized)
        if schema is None:
            schema = self._relaxng_schemas.get("*")
        if schema is None:
            return None
        return Path(schema)

    def _load_relaxng(
        self, schema_path: Path
    ) -> tuple[etree.RelaxNG | None, list[ValidationError]]:
        cached = self._relaxng_cache.get(schema_path)
        if cached is not None:
            return cached, []

        errors: list[ValidationError] = []
        if not schema_path.exists():
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=f"Relax NG schema file not found: {schema_path}",
                    part_uri=str(schema_path),
                    severity=ValidationSeverity.ERROR,
                )
            )
            return None, errors

        try:
            schema_doc = etree.parse(str(schema_path))
            relaxng = etree.RelaxNG(schema_doc)
            self._relaxng_cache[schema_path] = relaxng
            return relaxng, []
        except (etree.XMLSyntaxError, etree.RelaxNGParseError, OSError) as exc:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=f"Invalid Relax NG schema '{schema_path}': {exc}",
                    part_uri=str(schema_path),
                    severity=ValidationSeverity.ERROR,
                )
            )
            return None, errors

    @staticmethod
    def _expected_content_body_local(mimetype: str | None) -> str | None:
        if not mimetype:
            return None
        for prefix, local in OdfValidator.CONTENT_BODY_BY_MIMETYPE:
            if mimetype.startswith(prefix):
                return local
        return None

    @staticmethod
    def _first_element_child(element: etree._Element) -> etree._Element | None:
        for child in element:
            if isinstance(child.tag, str):
                return child
        return None

    def _validate_document_semantics(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        manifest_paths = package.manifest_paths()
        for part, expected_root in self.CORE_ROOTS.items():
            if part not in manifest_paths:
                continue
            root = parsed_parts.get(part)
            if root is None:
                continue

            qname = etree.QName(root)
            if qname.namespace != self.OFFICE_NS or qname.localname != expected_root:
                errors.append(
                    ValidationError(
                        error_type=ValidationErrorType.SCHEMA,
                        description=(
                            f"{part} root element must be office:{expected_root} "
                            f"(found '{root.tag}')"
                        ),
                        part_uri=self._normalize_part_uri(part),
                        severity=ValidationSeverity.ERROR,
                    )
                )

        content = parsed_parts.get("content.xml")
        if "content.xml" not in manifest_paths:
            return errors
        if content is None:
            return errors

        body = content.find(f"{{{self.OFFICE_NS}}}body")
        if body is None:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description="content.xml is missing required office:body element",
                    part_uri="/content.xml",
                    severity=ValidationSeverity.ERROR,
                )
            )
            return errors

        expected_body_local = self._expected_content_body_local(package.mimetype)
        if expected_body_local is None:
            return errors

        body_child = self._first_element_child(body)
        if body_child is None:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description="content.xml office:body has no document type element",
                    part_uri="/content.xml",
                    severity=ValidationSeverity.ERROR,
                )
            )
            return errors

        child_qname = etree.QName(body_child)
        if child_qname.namespace == self.OFFICE_NS and child_qname.localname == expected_body_local:
            return errors

        errors.append(
            ValidationError(
                error_type=ValidationErrorType.SEMANTIC,
                description=(
                    "content.xml body type does not match mimetype "
                    f"(expected office:{expected_body_local}, found '{body_child.tag}')"
                ),
                part_uri="/content.xml",
                severity=ValidationSeverity.ERROR,
            )
        )
        return errors

    def _validate_relaxng_parts(
        self,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        if not self._relaxng_validation:
            return []

        errors: list[ValidationError] = []
        for part, element in parsed_parts.items():
            schema_path = self._schema_path_for_part(part)
            if schema_path is None:
                continue

            relaxng, schema_errors = self._load_relaxng(schema_path)
            errors.extend(schema_errors)
            if relaxng is None:
                continue

            if relaxng.validate(element):
                continue

            detail = "Relax NG validation failed"
            if relaxng.error_log:
                last = relaxng.error_log.last_error
                if last is not None and last.message:
                    detail = f"Relax NG validation failed: {last.message}"
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=detail,
                    part_uri=self._normalize_part_uri(part),
                    severity=ValidationSeverity.ERROR,
                )
            )
        return errors

    def validate(self, path: str | Path) -> ValidationResult:
        errors: list[ValidationError] = []

        try:
            with OdfPackage(path) as package:
                errors.extend(package.validate_structure(strict=self._strict))
                parsed_parts, parse_errors = self._parse_xml_parts(package)
                errors.extend(parse_errors)
                errors.extend(self._validate_document_semantics(package, parsed_parts))
                errors.extend(self._validate_relaxng_parts(parsed_parts))
        except Exception as exc:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=str(exc),
                )
            )

        is_valid = not any(error.severity == ValidationSeverity.ERROR for error in errors)
        return ValidationResult(
            is_valid=is_valid,
            errors=errors,
            file_path=str(path),
            file_format=self._file_format,
        )
