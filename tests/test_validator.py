"""Tests for the main validator."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from openxml_audit import (
    FileFormat,
    OpenXmlValidator,
    ValidationResult,
    ValidationSeverity,
    is_valid_pptx,
    validate_pptx,
)
from openxml_audit.relationships import Relationship


class TestOpenXmlValidator:
    """Tests for OpenXmlValidator class."""

    def test_validator_initialization_defaults(self) -> None:
        """Test validator initializes with correct defaults."""
        validator = OpenXmlValidator()
        assert validator.file_format == FileFormat.OFFICE_2019
        assert validator.max_errors == 1000

    def test_validator_initialization_custom(self) -> None:
        """Test validator with custom options."""
        validator = OpenXmlValidator(
            file_format=FileFormat.OFFICE_2007,
            max_errors=50,
            schema_validation=False,
            semantic_validation=False,
        )
        assert validator.file_format == FileFormat.OFFICE_2007
        assert validator.max_errors == 50

    def test_validate_valid_pptx(self, minimal_pptx: Path) -> None:
        """Test validation of a valid PPTX file."""
        validator = OpenXmlValidator()
        result = validator.validate(minimal_pptx)

        assert isinstance(result, ValidationResult)
        assert result.file_path == str(minimal_pptx)
        assert result.file_format == FileFormat.OFFICE_2019
        # A well-formed minimal PPTX should be valid
        # (allowing for minor schema issues in our test fixture)

    def test_validate_returns_errors(
        self, invalid_pptx_missing_presentation: Path
    ) -> None:
        """Test validation returns errors for invalid PPTX."""
        validator = OpenXmlValidator()

        # This should raise PackageValidationError which is caught internally
        result = validator.validate(invalid_pptx_missing_presentation)

        assert not result.is_valid
        assert len(result.errors) > 0

    def test_validate_nonexistent_file(self, tmp_path: Path) -> None:
        """Test validation of nonexistent file."""
        validator = OpenXmlValidator()
        nonexistent = tmp_path / "nonexistent.pptx"

        # Should handle gracefully
        result = validator.validate(nonexistent)
        assert not result.is_valid

    def test_validate_not_a_zip(self, not_a_zip: Path) -> None:
        """Test validation of non-ZIP file."""
        validator = OpenXmlValidator()
        result = validator.validate(not_a_zip)

        assert not result.is_valid
        assert any("ZIP" in e.description for e in result.errors)

    def test_is_valid_method(self, minimal_pptx: Path) -> None:
        """Test the is_valid convenience method."""
        validator = OpenXmlValidator()
        # Should return boolean directly
        result = validator.is_valid(minimal_pptx)
        assert isinstance(result, bool)

    def test_max_errors_limit(self, minimal_pptx: Path) -> None:
        """Test that max_errors limits error collection."""
        validator = OpenXmlValidator(max_errors=1)
        result = validator.validate(minimal_pptx)

        # Should not exceed max_errors (counting only ERROR severity)
        error_count = sum(
            1 for e in result.errors if e.severity == ValidationSeverity.ERROR
        )
        assert error_count <= 1

    def test_max_errors_zero_unlimited(self, minimal_pptx: Path) -> None:
        """Test that max_errors=0 means unlimited."""
        validator = OpenXmlValidator(max_errors=0)
        # Should not crash with unlimited errors
        result = validator.validate(minimal_pptx)
        assert isinstance(result, ValidationResult)

    def test_validate_with_timings_returns_phase_metrics(self, minimal_pptx: Path) -> None:
        """Timing-enabled validation should return per-phase durations."""
        validator = OpenXmlValidator()
        result, timings = validator.validate_with_timings(minimal_pptx)

        assert isinstance(result, ValidationResult)
        expected_phases = {
            "package_structure",
            "profile_detection",
            "structure",
            "relationships",
            "binary",
            "schema",
            "semantic",
            "specific",
            "total",
        }
        assert set(timings.keys()) == expected_phases
        assert all(duration >= 0.0 for duration in timings.values())

    def test_validate_with_timings_schema_breakdown(self, minimal_pptx: Path) -> None:
        """Schema breakdown should be available when requested."""
        validator = OpenXmlValidator()
        _result, timings = validator.validate_with_timings(
            minimal_pptx,
            include_schema_breakdown=True,
        )

        expected_breakdown = {
            "schema.elements",
            "schema.constraint_lookup",
            "schema.children_expand",
            "schema.attributes",
            "schema.content_model",
            "schema.recursion",
        }
        assert expected_breakdown.issubset(timings.keys())

    def test_relationship_validation_ignores_fragment_only_targets(self) -> None:
        validator = OpenXmlValidator(schema_validation=False, semantic_validation=False)
        rel = Relationship(
            id="rId1",
            type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            target="#bookmark",
        )
        package = SimpleNamespace(
            relationships=[],
            list_parts=lambda: ["/word/document.xml"],
            get_part_relationships=(
                lambda source_uri: [rel]
                if source_uri == "/word/document.xml"
                else []
            ),
            has_part=lambda target_uri: target_uri == "/word/document.xml",
        )

        errors = validator._validate_relationships(package)  # type: ignore[arg-type]

        assert not errors


class TestValidationResult:
    """Tests for ValidationResult."""

    def test_error_count(self, minimal_pptx: Path) -> None:
        """Test error_count property."""
        validator = OpenXmlValidator()
        result = validator.validate(minimal_pptx)

        # error_count should match number of ERROR severity items
        expected = sum(
            1 for e in result.errors if e.severity == ValidationSeverity.ERROR
        )
        assert result.error_count == expected

    def test_warning_count(self, minimal_pptx: Path) -> None:
        """Test warning_count property."""
        validator = OpenXmlValidator()
        result = validator.validate(minimal_pptx)

        expected = sum(
            1 for e in result.errors if e.severity == ValidationSeverity.WARNING
        )
        assert result.warning_count == expected


class TestConvenienceFunctions:
    """Tests for convenience functions."""

    def test_validate_pptx(self, minimal_pptx: Path) -> None:
        """Test validate_pptx convenience function."""
        result = validate_pptx(minimal_pptx)
        assert isinstance(result, ValidationResult)

    def test_is_valid_pptx(self, minimal_pptx: Path) -> None:
        """Test is_valid_pptx convenience function."""
        result = is_valid_pptx(minimal_pptx)
        assert isinstance(result, bool)

    def test_validate_pptx_string_path(self, minimal_pptx: Path) -> None:
        """Test that string paths work."""
        result = validate_pptx(str(minimal_pptx))
        assert isinstance(result, ValidationResult)


class TestPresentationAppCompatParts:
    """Tests for the missing-presProps/viewProps/tableStyles check.

    PowerPoint triggers its "unreadable content" repair dialog when these
    parts are absent, even when the package is internally self-consistent
    (relationship and content-type entries removed too). The existing
    relationship-target check only catches the half-broken case.
    """

    REL_TYPES = {
        "presProps": (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
        ),
        "viewProps": (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
        ),
        "tableStyles": (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
        ),
    }
    PART_NAMES = {
        "presProps": "/ppt/presProps.xml",
        "viewProps": "/ppt/viewProps.xml",
        "tableStyles": "/ppt/tableStyles.xml",
    }

    def _strip(
        self, src: Path, dst: Path, labels: tuple[str, ...]
    ) -> None:
        """Copy `src` to `dst`, omitting the given app-compat parts and their
        content-type and relationship entries. Mirrors the issue's repro: the
        package becomes internally self-consistent."""
        import zipfile

        from lxml import etree

        rel_types = {self.REL_TYPES[k] for k in labels}
        part_names = {self.PART_NAMES[k] for k in labels}
        zip_paths = {n.lstrip("/") for n in part_names}

        with zipfile.ZipFile(src) as zin:
            files = {
                n: zin.read(n) for n in zin.namelist() if n not in zip_paths
            }

        ct_ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
        ct = etree.fromstring(files["[Content_Types].xml"])
        for override in ct.xpath("//ct:Override", namespaces=ct_ns):
            if override.get("PartName") in part_names:
                override.getparent().remove(override)
        files["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )

        rel_ns = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rels_path = "ppt/_rels/presentation.xml.rels"
        rels = etree.fromstring(files[rels_path])
        for rel in rels.xpath("//pr:Relationship", namespaces=rel_ns):
            if rel.get("Type") in rel_types:
                rel.getparent().remove(rel)
        files[rels_path] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)

    def test_minimal_pptx_passes_app_compat_check(
        self, minimal_pptx: Path
    ) -> None:
        """The minimal fixture ships with all three rels + parts present."""
        result = OpenXmlValidator().validate(minimal_pptx)
        descriptions = [e.description for e in result.errors]
        for label in ("presProps", "viewProps", "tableStyles"):
            assert not any(
                "missing" in d and label in d for d in descriptions
            ), descriptions

    def test_one_missing_app_compat_part_fires_error(
        self, minimal_pptx: Path, tmp_path: Path
    ) -> None:
        """Removing presProps (part + rel + content-type) → ERROR."""
        broken = tmp_path / "missing-presProps.pptx"
        self._strip(minimal_pptx, broken, ("presProps",))

        result = OpenXmlValidator().validate(broken)
        matching = [
            e for e in result.errors
            if "missing" in e.description and "presProps" in e.description
        ]
        assert matching, [e.description for e in result.errors]
        assert all(e.severity == ValidationSeverity.ERROR for e in matching)
        # Should not flag the other two
        assert not any(
            "missing" in e.description
            and ("viewProps" in e.description or "tableStyles" in e.description)
            for e in result.errors
        )

    def test_all_three_app_compat_parts_missing_fires_three_errors(
        self, minimal_pptx: Path, tmp_path: Path
    ) -> None:
        """The exact scenario from issue #4 — internally self-consistent file
        missing all three parts should flag is_valid=False with three errors."""
        broken = tmp_path / "missing-all-three.pptx"
        self._strip(minimal_pptx, broken, ("presProps", "viewProps", "tableStyles"))

        result = OpenXmlValidator().validate(broken)
        assert not result.is_valid
        for label in ("presProps", "viewProps", "tableStyles"):
            assert any(
                "missing" in e.description
                and label in e.description
                and e.severity == ValidationSeverity.ERROR
                for e in result.errors
            ), (label, [e.description for e in result.errors])

    def test_unrelated_check_still_catches_dangling_rel(
        self, minimal_pptx: Path, tmp_path: Path
    ) -> None:
        """If the rel still points at a missing part, the existing relationship-
        target check fires (independent of the new app-compat-parts check)."""
        import zipfile

        dst = tmp_path / "dangling-rel.pptx"
        with zipfile.ZipFile(minimal_pptx) as zin:
            files = {
                n: zin.read(n) for n in zin.namelist()
                if n != "ppt/presProps.xml"
            }
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)

        result = OpenXmlValidator().validate(dst)
        assert not result.is_valid
        # The rel still exists, so the new app-compat check should NOT fire
        # for presProps; the relationship-target check carries the load.
        new_check_hits = [
            e for e in result.errors
            if "missing presProps relationship" in e.description
        ]
        assert new_check_hits == [], [e.description for e in new_check_hits]
