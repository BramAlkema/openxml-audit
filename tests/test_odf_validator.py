"""Tests for ODF validator behavior."""

from __future__ import annotations

from pathlib import Path

import pytest

from openxml_audit.errors import FileFormat, ValidationErrorType, ValidationSeverity
from openxml_audit.odf import OdfValidator


class TestOdfValidator:
    """ODF end-to-end validator tests."""

    @staticmethod
    def _write_relaxng_schema(path: Path, *, root_name: str) -> None:
        path.write_text(
            f"""<?xml version="1.0" encoding="UTF-8"?>
<grammar xmlns="http://relaxng.org/ns/structure/1.0"
         xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <start>
    <element name="{root_name}">
      <zeroOrMore>
        <choice>
          <attribute>
            <anyName/>
          </attribute>
          <text/>
          <ref name="anyElement"/>
        </choice>
      </zeroOrMore>
    </element>
  </start>
  <define name="anyElement">
    <element>
      <anyName/>
      <zeroOrMore>
        <choice>
          <attribute>
            <anyName/>
          </attribute>
          <text/>
          <ref name="anyElement"/>
        </choice>
      </zeroOrMore>
    </element>
  </define>
</grammar>
""",
            encoding="utf-8",
        )

    def test_valid_odt_is_valid(self, minimal_odt: Path) -> None:
        result = OdfValidator().validate(minimal_odt)

        assert result.is_valid
        assert result.file_format == FileFormat.ODF_1_3

    def test_validator_initialization_defaults(self) -> None:
        validator = OdfValidator()
        assert validator.file_format == FileFormat.ODF_1_3
        assert validator.max_errors == 1000

    def test_validator_initialization_custom(self) -> None:
        validator = OdfValidator(
            file_format=FileFormat.ODF_1_2,
            max_errors=5,
            schema_validation=False,
            semantic_validation=False,
            strict=False,
        )
        assert validator.file_format == FileFormat.ODF_1_2
        assert validator.max_errors == 5

    def test_valid_ods_is_valid(self, minimal_ods: Path) -> None:
        result = OdfValidator(file_format=FileFormat.ODF_1_2).validate(minimal_ods)

        assert result.is_valid
        assert result.file_format == FileFormat.ODF_1_2

    def test_valid_odp_is_valid(self, minimal_odp: Path) -> None:
        result = OdfValidator().validate(minimal_odp)

        assert result.is_valid

    def test_broken_content_xml_reports_schema_error(self, odf_broken_content_xml: Path) -> None:
        result = OdfValidator().validate(odf_broken_content_xml)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and error.part_uri == "/content.xml"
            and "XML parse error" in error.description
            for error in result.errors
        )

    def test_unlisted_part_fails_in_strict_mode(self, odf_unlisted_xml_part: Path) -> None:
        result = OdfValidator(strict=True).validate(odf_unlisted_xml_part)

        assert not result.is_valid
        assert any(
            "is not declared in manifest.xml" in error.description
            and error.severity == ValidationSeverity.ERROR
            for error in result.errors
        )

    def test_unlisted_part_is_warning_in_permissive_mode(self, odf_unlisted_xml_part: Path) -> None:
        result = OdfValidator(strict=False).validate(odf_unlisted_xml_part)

        assert result.is_valid
        assert any(
            "is not declared in manifest.xml" in error.description
            and error.severity == ValidationSeverity.WARNING
            for error in result.errors
        )

    def test_is_valid_method(self, minimal_odt: Path) -> None:
        assert isinstance(OdfValidator().is_valid(minimal_odt), bool)

    def test_max_errors_limit(self, odf_invalid_mimetype: Path) -> None:
        result = OdfValidator(max_errors=1).validate(odf_invalid_mimetype)
        error_count = sum(
            1 for error in result.errors if error.severity == ValidationSeverity.ERROR
        )
        assert error_count <= 1

    def test_validate_with_timings_returns_phase_metrics(self, minimal_odt: Path) -> None:
        _result, timings = OdfValidator().validate_with_timings(minimal_odt)
        expected_phases = {
            "package_structure",
            "xml_parse",
            "schema",
            "semantic",
            "total",
        }
        assert set(timings.keys()) == expected_phases
        assert all(duration >= 0.0 for duration in timings.values())

    def test_content_body_type_mismatch_reports_semantic_error(
        self,
        odf_content_body_mismatch: Path,
    ) -> None:
        result = OdfValidator().validate(odf_content_body_mismatch)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SEMANTIC
            and "body type does not match mimetype" in error.description
            and error.part_uri == "/content.xml"
            for error in result.errors
        )

    def test_content_root_mismatch_reports_schema_error(
        self,
        odf_content_root_mismatch: Path,
    ) -> None:
        result = OdfValidator().validate(odf_content_root_mismatch)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and "content.xml root element must be office:document-content" in error.description
            and error.part_uri == "/content.xml"
            for error in result.errors
        )

    def test_relaxng_enabled_requires_schema_mapping(self) -> None:
        with pytest.raises(ValueError, match="relaxng_validation requires relaxng_schemas"):
            OdfValidator(relaxng_validation=True)

    def test_relaxng_requires_schema_validation_enabled(self) -> None:
        with pytest.raises(ValueError, match="relaxng_validation requires schema_validation=True"):
            OdfValidator(
                relaxng_validation=True,
                schema_validation=False,
                relaxng_schemas={"*": "x.rng"},
            )

    def test_relaxng_validation_passes_with_matching_schema(
        self,
        minimal_odt: Path,
        tmp_path: Path,
    ) -> None:
        schema_path = tmp_path / "content.rng"
        self._write_relaxng_schema(schema_path, root_name="office:document-content")

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": schema_path},
        ).validate(minimal_odt)

        assert result.is_valid
        assert not any("Relax NG validation failed" in error.description for error in result.errors)

    def test_relaxng_validation_reports_mismatch(
        self,
        minimal_odt: Path,
        tmp_path: Path,
    ) -> None:
        schema_path = tmp_path / "content.rng"
        self._write_relaxng_schema(schema_path, root_name="office:document-styles")

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": schema_path},
        ).validate(minimal_odt)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and error.part_uri == "/content.xml"
            and "Relax NG validation failed" in error.description
            for error in result.errors
        )

    def test_relaxng_missing_schema_file_reports_error(
        self,
        minimal_odt: Path,
        tmp_path: Path,
    ) -> None:
        missing = tmp_path / "missing.rng"
        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": missing},
        ).validate(minimal_odt)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.PACKAGE
            and "Relax NG schema file not found" in error.description
            for error in result.errors
        )
