"""Tests for M6 drawing, forms, and chart constraints."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.errors import ValidationErrorType, ValidationSeverity
from openxml_audit.odf import OdfValidator


class TestDrawingConstraints:
    """Drawing constraint tests (ODFSEMDRAW001-007)."""

    def test_draw_shape_missing_position_reports_warning(
        self, odf_draw_shape_no_position: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_shape_no_position)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW001" in e.id
            and "missing" in e.description.lower()
            and e.severity == ValidationSeverity.WARNING
            for e in result.errors
        )

    def test_draw_group_deep_nesting_reports_warning(
        self, odf_draw_group_deep_nesting: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_group_deep_nesting)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW002" in e.id
            and "nesting" in e.description.lower()
            and e.severity == ValidationSeverity.WARNING
            for e in result.errors
        )

    def test_draw_connector_unresolved_reports_error(
        self, odf_draw_connector_unresolved: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_connector_unresolved)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW003" in e.id
            and "nonexistent" in e.description.lower()
            for e in result.errors
        )

    def test_draw_custom_shape_no_geometry_reports_error(
        self, odf_draw_custom_shape_no_geometry: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_custom_shape_no_geometry)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW004" in e.id
            and "enhanced-geometry" in e.description.lower()
            for e in result.errors
        )

    def test_draw_frame_href_missing_part_reports_error(
        self, odf_draw_frame_href_missing_part: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_frame_href_missing_part)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW005" in e.id
            and "not present" in e.description.lower()
            for e in result.errors
        )

    def test_draw_3d_scene_empty_reports_warning(
        self, odf_draw_3d_scene_empty: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_3d_scene_empty)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW006" in e.id
            and "no 3D shape" in e.description
            and e.severity == ValidationSeverity.WARNING
            for e in result.errors
        )

    def test_draw_style_ref_unresolved_reports_error(
        self, odf_draw_style_ref_unresolved: Path
    ) -> None:
        result = OdfValidator().validate(odf_draw_style_ref_unresolved)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMDRAW007" in e.id
            and "NonExistentStyle" in e.description
            for e in result.errors
        )


class TestFormConstraints:
    """Form constraint tests (ODFSEMFORM001-004)."""

    def test_form_control_name_duplicate_reports_error(
        self, odf_form_control_name_duplicate: Path
    ) -> None:
        result = OdfValidator().validate(odf_form_control_name_duplicate)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMFORM001" in e.id
            and "Duplicate" in e.description
            and "Field1" in e.description
            for e in result.errors
        )

    def test_form_control_id_duplicate_reports_error(
        self, odf_form_control_id_duplicate: Path
    ) -> None:
        result = OdfValidator().validate(odf_form_control_id_duplicate)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMFORM002" in e.id
            and "dup1" in e.description
            for e in result.errors
        )

    def test_form_column_ref_unresolved_reports_error(
        self, odf_form_column_ref_unresolved: Path
    ) -> None:
        result = OdfValidator().validate(odf_form_column_ref_unresolved)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMFORM003" in e.id
            and "nonexistent_ctrl" in e.description
            for e in result.errors
        )

    def test_form_event_listener_no_href_reports_warning(
        self, odf_form_event_listener_no_href: Path
    ) -> None:
        result = OdfValidator().validate(odf_form_event_listener_no_href)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMFORM004" in e.id
            and "dom:click" in e.description
            and e.severity == ValidationSeverity.WARNING
            for e in result.errors
        )


class TestChartConstraints:
    """Chart constraint tests (ODFSEMCHART001-004)."""

    def test_chart_missing_plot_area_reports_error(
        self, odf_chart_missing_plot_area: Path
    ) -> None:
        result = OdfValidator().validate(odf_chart_missing_plot_area)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMCHART001" in e.id
            and "plot-area" in e.description.lower()
            for e in result.errors
        )

    def test_chart_no_axis_reports_warning(
        self, odf_chart_no_axis: Path
    ) -> None:
        result = OdfValidator().validate(odf_chart_no_axis)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMCHART002" in e.id
            and "no chart:axis" in e.description
            and e.severity == ValidationSeverity.WARNING
            for e in result.errors
        )

    def test_chart_empty_range_address_reports_error(
        self, odf_chart_empty_range_address: Path
    ) -> None:
        result = OdfValidator().validate(odf_chart_empty_range_address)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMCHART003" in e.id
            and "empty" in e.description.lower()
            for e in result.errors
        )

    def test_chart_style_ref_unresolved_reports_error(
        self, odf_chart_style_ref_unresolved: Path
    ) -> None:
        result = OdfValidator().validate(odf_chart_style_ref_unresolved)
        assert any(
            e.error_type == ValidationErrorType.SEMANTIC
            and "ODFSEMCHART004" in e.id
            and "NonExistentChartStyle" in e.description
            for e in result.errors
        )


class TestM6Regression:
    """Regression tests: valid files must not trigger M6 rules."""

    def test_valid_odt_no_drawing_errors(self, minimal_odt: Path) -> None:
        result = OdfValidator().validate(minimal_odt)
        m6_ids = {"ODFSEMDRAW", "ODFSEMFORM", "ODFSEMCHART"}
        assert not any(
            any(e.id.startswith(prefix) for prefix in m6_ids)
            for e in result.errors
        )

    def test_valid_ods_no_m6_errors(self, minimal_ods: Path) -> None:
        result = OdfValidator(semantic_validation=True).validate(minimal_ods)
        m6_ids = {"ODFSEMDRAW", "ODFSEMFORM", "ODFSEMCHART"}
        assert not any(
            any(e.id.startswith(prefix) for prefix in m6_ids)
            for e in result.errors
        )
