"""Tests for M5 bundled Relax NG schema validation."""

from __future__ import annotations

from pathlib import Path

from openxml_audit.odf import OdfValidator
from openxml_audit.odf.schema_core import OdfRelaxNgRouter
from openxml_audit.odf.schemas import (
    BUNDLED_VERSIONS,
    build_bundled_routes,
    get_bundled_schema_dir,
    get_bundled_schema_path,
)

# ── Schema package tests ─────────────────────────────────────────────────


def test_bundled_versions() -> None:
    assert "1.2" in BUNDLED_VERSIONS
    assert "1.3" in BUNDLED_VERSIONS


def test_get_bundled_schema_dir_exists() -> None:
    for version in BUNDLED_VERSIONS:
        schema_dir = get_bundled_schema_dir(version)
        assert schema_dir is not None
        assert schema_dir.is_dir()


def test_get_bundled_schema_dir_unknown_version() -> None:
    assert get_bundled_schema_dir("9.9") is None


def test_get_bundled_schema_path_content() -> None:
    for version in BUNDLED_VERSIONS:
        path = get_bundled_schema_path("content.xml", version)
        assert path is not None
        assert path.exists()
        assert path.suffix == ".rng"


def test_get_bundled_schema_path_all_parts() -> None:
    for version in BUNDLED_VERSIONS:
        for part in ("content.xml", "styles.xml", "meta.xml", "settings.xml"):
            path = get_bundled_schema_path(part, version)
            assert path is not None, f"Missing schema for {part} at {version}"


def test_get_bundled_schema_path_manifest() -> None:
    for version in BUNDLED_VERSIONS:
        path = get_bundled_schema_path("META-INF/manifest.xml", version)
        assert path is not None, f"Missing manifest schema at {version}"


def test_get_bundled_schema_path_unknown_part() -> None:
    assert get_bundled_schema_path("unknown.xml", "1.3") is None


def test_build_bundled_routes() -> None:
    routes = build_bundled_routes()
    assert "1.2" in routes
    assert "1.3" in routes
    assert "content.xml" in routes["1.3"]
    assert "META-INF/manifest.xml" in routes["1.3"]


def test_router_from_bundled() -> None:
    router = OdfRelaxNgRouter.from_bundled()
    assert not router.is_empty()
    route = router.resolve("content.xml", "1.3")
    assert route is not None
    assert route.schema_path.exists()


def test_router_from_bundled_resolves_manifest() -> None:
    router = OdfRelaxNgRouter.from_bundled()
    route = router.resolve("META-INF/manifest.xml", "1.3")
    assert route is not None
    assert "manifest" in route.schema_path.name.lower()


# ── Zero-config validation tests ─────────────────────────────────────────


def test_zero_config_relaxng_valid_odt(minimal_odt: Path) -> None:
    """OdfValidator(relaxng_validation=True) works without explicit schemas."""
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(minimal_odt)
    schema_errors = [e for e in result.errors if "Relax NG validation failed" in e.description]
    assert not schema_errors


def test_zero_config_relaxng_valid_ods(minimal_ods: Path) -> None:
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(minimal_ods)
    schema_errors = [e for e in result.errors if "Relax NG validation failed" in e.description]
    assert not schema_errors


def test_zero_config_relaxng_valid_odp(minimal_odp: Path) -> None:
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(minimal_odp)
    schema_errors = [e for e in result.errors if "Relax NG validation failed" in e.description]
    assert not schema_errors


def test_zero_config_relaxng_valid_v12(minimal_odt_v12: Path) -> None:
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(minimal_odt_v12)
    schema_errors = [e for e in result.errors if "Relax NG validation failed" in e.description]
    assert not schema_errors


def test_zero_config_relaxng_detects_invalid_structure(
    odf_content_body_mismatch: Path,
) -> None:
    """Bundled schema detects structural errors in document content."""
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(odf_content_body_mismatch)
    # The OASIS schema validates full document structure;
    # a body type mismatch or invalid child element causes schema errors
    has_errors = not result.is_valid
    assert has_errors


def test_zero_config_manifest_schema_valid(minimal_odt: Path) -> None:
    """Manifest schema validates correctly for a valid package."""
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(minimal_odt)
    manifest_errors = [
        e
        for e in result.errors
        if e.part_uri == "/META-INF/manifest.xml" and "Relax NG validation failed" in e.description
    ]
    assert not manifest_errors


def test_zero_config_with_styles(minimal_odt_with_styles: Path) -> None:
    """Both content.xml and styles.xml validate with bundled schemas."""
    result = OdfValidator(
        relaxng_validation=True,
        require_schema_routes=False,
    ).validate(minimal_odt_with_styles)
    schema_errors = [e for e in result.errors if "Relax NG validation failed" in e.description]
    assert not schema_errors
