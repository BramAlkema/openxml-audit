"""ODF validator aligned with OOXML validator pipeline shape."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from time import perf_counter
from typing import cast

from lxml import etree

from openxml_audit.errors import (
    FileFormat,
    PackageValidationError,
    ValidationError,
    ValidationErrorType,
    ValidationResult,
    ValidationSeverity,
)
from openxml_audit.odf._helpers import normalize_part_uri
from openxml_audit.odf.package import OdfPackage
from openxml_audit.odf.schema_core import (
    OdfRelaxNgResolver,
    OdfRelaxNgRouter,
    detect_odf_schema_version,
)
from openxml_audit.odf.security import OdfCryptographicVerifier, OdfSecurityValidator
from openxml_audit.odf.semantic import OdfSemanticValidator


class OdfValidator:
    """Validate ODF packages using package, schema, and semantic phases."""

    def __init__(
        self,
        file_format: FileFormat = FileFormat.ODF_1_3,
        max_errors: int = 1000,
        schema_validation: bool = True,
        semantic_validation: bool = True,
        security_validation: bool = False,
        strict: bool = True,
        *,
        relaxng_validation: bool = False,
        relaxng_schemas: Mapping[str, str | Path] | None = None,
        schema_routes: Mapping[str, Mapping[str, str | Path]] | None = None,
        require_schema_routes: bool = True,
        max_schema_parts: int = 128,
        max_schema_xml_bytes: int = 8_000_000,
        verify_cryptography: bool = False,
        cryptographic_verifier: OdfCryptographicVerifier | None = None,
    ):
        if max_schema_parts < 0:
            raise ValueError("max_schema_parts must be >= 0")
        if max_schema_xml_bytes < 0:
            raise ValueError("max_schema_xml_bytes must be >= 0")
        if relaxng_validation and not schema_validation:
            raise ValueError("relaxng_validation requires schema_validation=True")

        self._file_format = file_format
        self._max_errors = max_errors
        self._schema_validation = schema_validation
        self._semantic_validation = semantic_validation
        self._security_validation = security_validation
        self._strict = strict
        self._relaxng_validation = relaxng_validation and schema_validation
        self._schema_require_routes = require_schema_routes
        self._max_schema_parts = max_schema_parts
        self._max_schema_xml_bytes = max_schema_xml_bytes
        if schema_routes is not None:
            self._schema_router = OdfRelaxNgRouter(schema_routes)
        elif relaxng_schemas is not None:
            self._schema_router = OdfRelaxNgRouter.from_legacy_mapping(relaxng_schemas)
        elif relaxng_validation:
            # Use bundled schemas for zero-config validation
            self._schema_router = OdfRelaxNgRouter.from_bundled()
        else:
            self._schema_router = OdfRelaxNgRouter()
        self._schema_resolver = OdfRelaxNgResolver()
        self._semantic_core_validator = OdfSemanticValidator()
        self._security_core_validator = (
            OdfSecurityValidator(
                verify_cryptography=verify_cryptography,
                cryptographic_verifier=cryptographic_verifier,
            )
            if security_validation
            else None
        )
        self._relaxng_cache: dict[Path, etree.RelaxNG] = {}

    @property
    def file_format(self) -> FileFormat:
        """Get the target ODF version."""
        return self._file_format

    @property
    def max_errors(self) -> int:
        """Get the maximum number of ERROR entries to collect."""
        return self._max_errors

    @staticmethod
    def _normalize_part_uri(part_path: str) -> str:
        return normalize_part_uri(part_path)

    def validate(self, path: str | Path) -> ValidationResult:
        """Validate an ODF file."""
        result, _timings = self.validate_with_timings(path)
        return result

    def validate_with_timings(
        self,
        path: str | Path,
        include_schema_breakdown: bool = False,
    ) -> tuple[ValidationResult, dict[str, float]]:
        """Validate an ODF file and return per-phase timing metrics."""
        path = Path(path)
        errors: list[ValidationError] = []
        timings: dict[str, float] = {
            "package_structure": 0.0,
            "xml_parse": 0.0,
            "schema": 0.0,
            "semantic": 0.0,
            "security": 0.0,
            "total": 0.0,
        }
        total_start = perf_counter()

        def finish() -> tuple[ValidationResult, dict[str, float]]:
            timings["total"] = perf_counter() - total_start
            return self._create_result(path, errors), timings

        parsed_parts: dict[str, etree._Element] = {}
        parsed_part_sizes: dict[str, int] = {}
        try:
            with OdfPackage(path) as package_handle:
                package = cast(OdfPackage, package_handle)
                phase_start = perf_counter()
                errors.extend(self._validate_package_structure(package))
                timings["package_structure"] += perf_counter() - phase_start
                if self._trim_and_check_limit(errors):
                    return finish()

                if (
                    self._schema_validation
                    or self._semantic_validation
                    or self._security_validation
                ):
                    phase_start = perf_counter()
                    parsed_parts, parsed_part_sizes, parse_errors = self._parse_xml_parts(package)
                    errors.extend(parse_errors)
                    timings["xml_parse"] += perf_counter() - phase_start
                    if self._trim_and_check_limit(errors):
                        return finish()

                if self._schema_validation:
                    phase_start = perf_counter()
                    schema_errors = self._validate_schema(
                        package,
                        parsed_parts,
                        parsed_part_sizes,
                    )
                    errors.extend(schema_errors)
                    timings["schema"] += perf_counter() - phase_start
                    if include_schema_breakdown and self._relaxng_validation:
                        timings["schema.relaxng"] = timings["schema"]
                    if self._trim_and_check_limit(errors):
                        return finish()

                if self._semantic_validation:
                    phase_start = perf_counter()
                    errors.extend(self._validate_document_semantics(package, parsed_parts))
                    timings["semantic"] += perf_counter() - phase_start
                    if self._trim_and_check_limit(errors):
                        return finish()

                if self._security_core_validator is not None:
                    phase_start = perf_counter()
                    errors.extend(self._security_core_validator.validate(package, parsed_parts))
                    self._trim_and_check_limit(errors)
                    timings["security"] += perf_counter() - phase_start

        except PackageValidationError as exc:
            errors.extend(exc.errors)
            self._trim_and_check_limit(errors)
        except Exception as exc:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.PACKAGE,
                    description=str(exc),
                )
            )

        return finish()

    def is_valid(self, path: str | Path) -> bool:
        """Quick check if an ODF file is valid."""
        return self.validate(path).is_valid

    def _trim_and_check_limit(self, errors: list[ValidationError]) -> bool:
        """Trim errors to limit and return True if the limit has been reached."""
        if self._max_errors == 0:
            return False
        limited: list[ValidationError] = []
        error_count = 0
        for error in errors:
            if error.severity == ValidationSeverity.ERROR:
                if error_count >= self._max_errors:
                    continue
                error_count += 1
            limited.append(error)
        if len(limited) != len(errors):
            errors[:] = limited
        return error_count >= self._max_errors

    def _create_result(self, path: Path, errors: list[ValidationError]) -> ValidationResult:
        is_valid = not any(error.severity == ValidationSeverity.ERROR for error in errors)
        return ValidationResult(
            is_valid=is_valid,
            errors=errors,
            file_path=str(path),
            file_format=self._file_format,
        )

    def _validate_package_structure(self, package: OdfPackage) -> list[ValidationError]:
        return package.validate_structure(strict=self._strict)

    def _collect_xml_parts(self, package: OdfPackage) -> list[str]:
        parts: set[str] = set(package.list_xml_parts())
        for core in ("content.xml", "styles.xml", "meta.xml", "settings.xml"):
            if package.has_part(core):
                parts.add(core)
        return sorted(parts)

    def _parse_xml_parts(
        self,
        package: OdfPackage,
    ) -> tuple[dict[str, etree._Element], dict[str, int], list[ValidationError]]:
        parsed_parts: dict[str, etree._Element] = {}
        parsed_part_sizes: dict[str, int] = {}
        errors: list[ValidationError] = []
        for part in self._collect_xml_parts(package):
            content = package.get_part_content(part)
            if content is None:
                continue
            parsed_part_sizes[part] = len(content)
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
        return parsed_parts, parsed_part_sizes, errors

    def _validate_schema(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
        parsed_part_sizes: dict[str, int],
    ) -> list[ValidationError]:
        if not self._relaxng_validation:
            return []
        guardrail_errors = self._validate_schema_guardrails(parsed_parts, parsed_part_sizes)
        if guardrail_errors:
            return guardrail_errors

        schema_version = detect_odf_schema_version(
            package,
            parsed_parts,
            fallback_format=self._file_format,
        )
        errors = self._validate_relaxng_parts(parsed_parts, schema_version=schema_version)
        errors.extend(self._validate_manifest_schema(package, schema_version))
        return errors

    def _validate_schema_guardrails(
        self,
        parsed_parts: Mapping[str, etree._Element],
        parsed_part_sizes: Mapping[str, int],
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        if self._max_schema_parts > 0 and len(parsed_parts) > self._max_schema_parts:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=(
                        "Schema-core part guardrail exceeded: "
                        f"{len(parsed_parts)} XML parts exceeds max_schema_parts="
                        f"{self._max_schema_parts}"
                    ),
                    part_uri="/",
                    severity=ValidationSeverity.ERROR,
                    id="ODFSCHEMA003",
                )
            )

        total_bytes = sum(parsed_part_sizes.values())
        if self._max_schema_xml_bytes > 0 and total_bytes > self._max_schema_xml_bytes:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=(
                        "Schema-core XML-size guardrail exceeded: "
                        f"{total_bytes} bytes exceeds max_schema_xml_bytes="
                        f"{self._max_schema_xml_bytes}"
                    ),
                    part_uri="/",
                    severity=ValidationSeverity.ERROR,
                    id="ODFSCHEMA004",
                )
            )
        return errors

    def _schema_path_for_part(self, part: str, schema_version: str) -> Path | None:
        route = self._schema_router.resolve(part, schema_version)
        if route is None:
            return None
        return route.schema_path

    def _load_relaxng(
        self, schema_path: Path
    ) -> tuple[etree.RelaxNG | None, list[ValidationError]]:
        normalized_path = schema_path.resolve()
        cached = self._relaxng_cache.get(normalized_path)
        if cached is not None:
            return cached, []

        errors: list[ValidationError] = []
        for reference_error in self._schema_resolver.preflight_references(normalized_path):
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=reference_error,
                    part_uri=str(normalized_path),
                    severity=ValidationSeverity.ERROR,
                    id="ODFSCHEMA002",
                )
            )
        if errors:
            return None, errors

        try:
            schema_doc = self._schema_resolver.parse_schema(normalized_path)
            relaxng = etree.RelaxNG(schema_doc)
            self._relaxng_cache[normalized_path] = relaxng
            return relaxng, []
        except (etree.XMLSyntaxError, etree.RelaxNGParseError, OSError) as exc:
            errors.append(
                ValidationError(
                    error_type=ValidationErrorType.SCHEMA,
                    description=f"Invalid Relax NG schema '{normalized_path}': {exc}",
                    part_uri=str(normalized_path),
                    severity=ValidationSeverity.ERROR,
                    id="ODFSCHEMA002",
                )
            )
            return None, errors

    def _validate_document_semantics(
        self,
        package: OdfPackage,
        parsed_parts: dict[str, etree._Element],
    ) -> list[ValidationError]:
        return self._semantic_core_validator.validate(package, parsed_parts)

    def _validate_relaxng_parts(
        self,
        parsed_parts: dict[str, etree._Element],
        *,
        schema_version: str,
    ) -> list[ValidationError]:
        errors: list[ValidationError] = []
        for part in sorted(parsed_parts):
            element = parsed_parts[part]
            schema_path = self._schema_path_for_part(part, schema_version)
            if schema_path is None:
                if self._schema_require_routes:
                    errors.append(
                        ValidationError(
                            error_type=ValidationErrorType.SCHEMA,
                            description=(
                                "No Relax NG schema route for manifest XML part "
                                f"'{self._normalize_part_uri(part)}' at ODF version "
                                f"'{schema_version}'"
                            ),
                            part_uri=self._normalize_part_uri(part),
                            severity=ValidationSeverity.ERROR,
                            id="ODFSCHEMA001",
                        )
                    )
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

    def _validate_manifest_schema(
        self,
        package: OdfPackage,
        schema_version: str,
    ) -> list[ValidationError]:
        """Validate META-INF/manifest.xml against its Relax NG schema if available."""
        manifest_part = "META-INF/manifest.xml"
        schema_path = self._schema_path_for_part(manifest_part, schema_version)
        if schema_path is None:
            return []

        content = package.get_part_content(manifest_part)
        if content is None:
            return []

        try:
            manifest_xml = etree.fromstring(content)
        except etree.XMLSyntaxError:
            return []  # Already reported during XML parse phase

        relaxng, schema_errors = self._load_relaxng(schema_path)
        if schema_errors:
            return schema_errors
        if relaxng is None:
            return []

        if relaxng.validate(manifest_xml):
            return []

        detail = "Relax NG validation failed"
        if relaxng.error_log:
            last = relaxng.error_log.last_error
            if last is not None and last.message:
                detail = f"Relax NG validation failed: {last.message}"
        return [
            ValidationError(
                error_type=ValidationErrorType.SCHEMA,
                description=detail,
                part_uri=self._normalize_part_uri(manifest_part),
                severity=ValidationSeverity.ERROR,
            )
        ]
