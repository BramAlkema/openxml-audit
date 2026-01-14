"""Tests for relationship parsing and handling."""

from __future__ import annotations

import pytest

from openxml_audit.relationships import Relationship, RelationshipCollection, get_rels_path
from tests.fixture_loader import load_fixture_bytes


class TestRelationship:
    """Tests for Relationship dataclass."""

    def test_is_external_http(self) -> None:
        """Test external relationship detection for HTTP."""
        rel = Relationship(
            id="rId1",
            type="http://example.com/type",
            target="http://example.com/external",
            target_mode="External",
        )
        assert rel.is_external

    def test_is_internal(self) -> None:
        """Test internal relationship detection."""
        rel = Relationship(
            id="rId1",
            type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            target="slides/slide1.xml",
            target_mode=None,
        )
        assert not rel.is_external

    def test_resolve_target_absolute(self) -> None:
        """Test resolving absolute target path."""
        rel = Relationship(
            id="rId1",
            type="http://example.com/type",
            target="/ppt/slides/slide1.xml",
        )
        resolved = rel.resolve_target("/ppt/presentation.xml")
        assert resolved == "/ppt/slides/slide1.xml"

    def test_resolve_target_relative(self) -> None:
        """Test resolving relative target path."""
        rel = Relationship(
            id="rId1",
            type="http://example.com/type",
            target="slides/slide1.xml",
        )
        resolved = rel.resolve_target("/ppt/presentation.xml")
        assert resolved == "/ppt/slides/slide1.xml"

    def test_resolve_target_with_parent(self) -> None:
        """Test resolving target with parent directory reference."""
        rel = Relationship(
            id="rId1",
            type="http://example.com/type",
            target="../theme/theme1.xml",
        )
        resolved = rel.resolve_target("/ppt/slideMasters/slideMaster1.xml")
        assert resolved == "/ppt/theme/theme1.xml"


class TestRelationshipCollection:
    """Tests for RelationshipCollection."""

    def test_from_xml(self) -> None:
        """Test parsing relationships from XML."""
        xml = load_fixture_bytes("relationships", "two_rels.xml")

        rels = RelationshipCollection.from_xml(xml)
        assert len(rels) == 2

    def test_get_by_id(self) -> None:
        """Test getting relationship by ID."""
        xml = load_fixture_bytes("relationships", "by_id.xml")

        rels = RelationshipCollection.from_xml(xml)

        rel1 = rels.get_by_id("rId1")
        assert rel1 is not None
        assert rel1.target == "target1.xml"

        rel2 = rels.get_by_id("rId2")
        assert rel2 is not None
        assert rel2.target == "target2.xml"

        assert rels.get_by_id("rId999") is None

    def test_get_by_type(self) -> None:
        """Test getting relationships by type."""
        xml = load_fixture_bytes("relationships", "by_type.xml")

        rels = RelationshipCollection.from_xml(xml)

        slide_rels = list(rels.get_by_type(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
        ))
        assert len(slide_rels) == 2

        master_rels = list(rels.get_by_type(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
        ))
        assert len(master_rels) == 1

    def test_resolve_target(self) -> None:
        """Test resolve_target helper method."""
        xml = load_fixture_bytes("relationships", "resolve_target.xml")

        # Source URI needs to be passed at construction time
        rels = RelationshipCollection.from_xml(xml, source_uri="/ppt/presentation.xml")

        resolved = rels.resolve_target("rId1")
        assert resolved == "/ppt/slides/slide1.xml"

        # Non-existent ID
        assert rels.resolve_target("rId999") is None

    def test_iteration(self) -> None:
        """Test iterating over relationships."""
        xml = load_fixture_bytes("relationships", "iteration.xml")

        rels = RelationshipCollection.from_xml(xml)

        ids = [r.id for r in rels]
        assert "rId1" in ids
        assert "rId2" in ids

    def test_empty_relationships(self) -> None:
        """Test parsing empty relationships XML."""
        xml = load_fixture_bytes("relationships", "empty.xml")

        rels = RelationshipCollection.from_xml(xml)
        assert len(rels) == 0


class TestGetRelsPath:
    """Tests for get_rels_path function."""

    def test_root_part(self) -> None:
        """Test getting rels path for root part."""
        assert get_rels_path("/") == "/_rels/.rels"

    def test_presentation_part(self) -> None:
        """Test getting rels path for presentation.xml."""
        assert get_rels_path("/ppt/presentation.xml") == "/ppt/_rels/presentation.xml.rels"

    def test_slide_part(self) -> None:
        """Test getting rels path for slide."""
        assert get_rels_path("/ppt/slides/slide1.xml") == "/ppt/slides/_rels/slide1.xml.rels"

    def test_without_leading_slash(self) -> None:
        """Test path without leading slash."""
        result = get_rels_path("ppt/presentation.xml")
        assert "_rels/presentation.xml.rels" in result
