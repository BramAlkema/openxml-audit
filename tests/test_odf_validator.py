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

    @staticmethod
    def _write_relaxng_any_root_schema(path: Path) -> None:
        path.write_text(
            """<?xml version="1.0" encoding="UTF-8"?>
<grammar xmlns="http://relaxng.org/ns/structure/1.0">
  <start>
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

    @staticmethod
    def _write_relaxng_schema_with_include(
        path: Path,
        *,
        include_filename: str,
    ) -> None:
        path.write_text(
            f"""<?xml version="1.0" encoding="UTF-8"?>
<grammar xmlns="http://relaxng.org/ns/structure/1.0">
  <include href="{include_filename}"/>
  <start>
    <ref name="documentRoot"/>
  </start>
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

    def test_valid_odt_v12_markers_are_accepted(self, minimal_odt_v12: Path) -> None:
        result = OdfValidator(file_format=FileFormat.ODF_1_2).validate(minimal_odt_v12)
        assert result.is_valid

    def test_valid_odt_v14_markers_are_accepted(self, minimal_odt_v14: Path) -> None:
        result = OdfValidator().validate(minimal_odt_v14)
        assert result.is_valid

    def test_signed_stub_package_is_accepted_at_foundation_level(
        self,
        minimal_odt_signed_stub: Path,
    ) -> None:
        result = OdfValidator().validate(minimal_odt_signed_stub)
        assert result.is_valid

    def test_encrypted_stub_package_is_accepted_at_foundation_level(
        self,
        minimal_odt_encrypted_stub: Path,
    ) -> None:
        result = OdfValidator().validate(minimal_odt_encrypted_stub)
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
            "security",
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

    def test_auxiliary_invalid_styles_xml_reports_schema_error(
        self,
        odf_aux_invalid_styles_xml: Path,
    ) -> None:
        result = OdfValidator().validate(odf_aux_invalid_styles_xml)
        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and error.part_uri == "/styles.xml"
            and "XML parse error" in error.description
            for error in result.errors
        )

    def test_relaxng_enabled_uses_bundled_schemas_by_default(self) -> None:
        validator = OdfValidator(relaxng_validation=True)
        assert not validator._schema_router.is_empty()

    def test_relaxng_requires_schema_validation_enabled(self) -> None:
        with pytest.raises(ValueError, match="relaxng_validation requires schema_validation=True"):
            OdfValidator(
                relaxng_validation=True,
                schema_validation=False,
                relaxng_schemas={"*": "x.rng"},
            )

    def test_schema_guardrail_arguments_must_be_non_negative(self) -> None:
        with pytest.raises(ValueError, match="max_schema_parts must be >= 0"):
            OdfValidator(relaxng_validation=False, max_schema_parts=-1)
        with pytest.raises(ValueError, match="max_schema_xml_bytes must be >= 0"):
            OdfValidator(relaxng_validation=False, max_schema_xml_bytes=-1)

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
            error.error_type == ValidationErrorType.SCHEMA
            and "No such file or directory" in error.description
            for error in result.errors
        )

    @pytest.mark.parametrize("fixture_name", ["minimal_ods", "minimal_odp"])
    def test_relaxng_validation_accepts_ods_and_odp_with_any_root_schema(
        self,
        request: pytest.FixtureRequest,
        fixture_name: str,
        tmp_path: Path,
    ) -> None:
        package = request.getfixturevalue(fixture_name)
        schema_path = tmp_path / "any_root.rng"
        self._write_relaxng_any_root_schema(schema_path)

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": schema_path},
            require_schema_routes=False,
        ).validate(package)

        assert result.is_valid

    def test_schema_core_reports_missing_route_for_manifest_declared_xml_member(
        self,
        minimal_odt_with_styles: Path,
        tmp_path: Path,
    ) -> None:
        content_schema = tmp_path / "content.rng"
        self._write_relaxng_schema(content_schema, root_name="office:document-content")

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": content_schema},
        ).validate(minimal_odt_with_styles)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and error.part_uri == "/styles.xml"
            and "No Relax NG schema route" in error.description
            for error in result.errors
        )

    def test_schema_core_validates_all_manifest_declared_xml_members_when_routes_exist(
        self,
        minimal_odt_with_styles: Path,
        tmp_path: Path,
    ) -> None:
        content_schema = tmp_path / "content.rng"
        styles_schema = tmp_path / "styles.rng"
        self._write_relaxng_schema(content_schema, root_name="office:document-content")
        self._write_relaxng_schema(styles_schema, root_name="office:document-styles")

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={
                "content.xml": content_schema,
                "styles.xml": styles_schema,
            },
        ).validate(minimal_odt_with_styles)

        assert result.is_valid

    def test_schema_core_routes_14_packages_with_version_specific_mapping(
        self,
        minimal_odt_v14: Path,
        tmp_path: Path,
    ) -> None:
        schema_v13 = tmp_path / "content_13.rng"
        schema_v14 = tmp_path / "content_14.rng"
        self._write_relaxng_schema(schema_v13, root_name="office:document-styles")
        self._write_relaxng_schema(schema_v14, root_name="office:document-content")

        result = OdfValidator(
            relaxng_validation=True,
            schema_routes={
                "1.3": {"content.xml": schema_v13},
                "1.4": {"content.xml": schema_v14},
            },
        ).validate(minimal_odt_v14)

        assert result.is_valid

    def test_schema_core_relaxng_resolver_supports_include_references(
        self,
        minimal_odt: Path,
        tmp_path: Path,
    ) -> None:
        main_schema = tmp_path / "main.rng"
        shared_schema = tmp_path / "shared.rng"
        self._write_relaxng_schema_with_include(main_schema, include_filename="shared.rng")
        shared_schema.write_text(
            """<?xml version="1.0" encoding="UTF-8"?>
<grammar xmlns="http://relaxng.org/ns/structure/1.0"
         xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <define name="documentRoot">
    <element name="office:document-content">
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

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": main_schema},
        ).validate(minimal_odt)

        assert result.is_valid

    def test_schema_core_relaxng_resolver_reports_missing_include_reference(
        self,
        minimal_odt: Path,
        tmp_path: Path,
    ) -> None:
        main_schema = tmp_path / "main.rng"
        self._write_relaxng_schema_with_include(main_schema, include_filename="missing.rng")

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"content.xml": main_schema},
        ).validate(minimal_odt)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and "Unresolvable Relax NG reference" in error.description
            for error in result.errors
        )

    def test_schema_core_part_guardrail_reports_error(
        self,
        minimal_odt_with_styles: Path,
        tmp_path: Path,
    ) -> None:
        any_root_schema = tmp_path / "any_root.rng"
        self._write_relaxng_any_root_schema(any_root_schema)

        result = OdfValidator(
            relaxng_validation=True,
            relaxng_schemas={"*": any_root_schema},
            max_schema_parts=1,
        ).validate(minimal_odt_with_styles)

        assert not result.is_valid
        assert any(
            error.error_type == ValidationErrorType.SCHEMA
            and "Schema-core part guardrail exceeded" in error.description
            for error in result.errors
        )
