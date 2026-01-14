"""Tests for schema type validators."""

from __future__ import annotations

import pytest

from openxml_audit.schema.types import (
    AnyURITypeValidator,
    BooleanTypeValidator,
    DateTimeTypeValidator,
    DecimalTypeValidator,
    HexBinaryTypeValidator,
    IntegerTypeValidator,
    NCNameTypeValidator,
    StringTypeValidator,
)


class TestStringTypeValidator:
    """Tests for StringTypeValidator."""

    def test_valid_string(self) -> None:
        """Test validation of valid string."""
        validator = StringTypeValidator()
        result = validator.validate("hello")
        assert result.is_valid

    def test_min_length(self) -> None:
        """Test min length constraint."""
        validator = StringTypeValidator(min_length=3)

        result = validator.validate("abc")
        assert result.is_valid

        result = validator.validate("ab")
        assert not result.is_valid
        assert "minimum" in result.error_message.lower() or "less than" in result.error_message.lower()

    def test_max_length(self) -> None:
        """Test max length constraint."""
        validator = StringTypeValidator(max_length=5)

        result = validator.validate("hello")
        assert result.is_valid

        result = validator.validate("hello!")
        assert not result.is_valid
        assert "maximum" in result.error_message.lower() or "exceeds" in result.error_message.lower()

    def test_pattern(self) -> None:
        """Test pattern constraint."""
        validator = StringTypeValidator(pattern=r"^\d{3}-\d{2}-\d{4}$")

        result = validator.validate("123-45-6789")
        assert result.is_valid

        result = validator.validate("invalid")
        assert not result.is_valid
        assert "pattern" in result.error_message.lower()

    def test_enumeration(self) -> None:
        """Test enumeration constraint."""
        validator = StringTypeValidator(enumeration=["red", "green", "blue"])

        result = validator.validate("red")
        assert result.is_valid

        result = validator.validate("green")
        assert result.is_valid

        result = validator.validate("yellow")
        assert not result.is_valid
        assert "allowed values" in result.error_message.lower() or "not in" in result.error_message.lower()


class TestBooleanTypeValidator:
    """Tests for BooleanTypeValidator."""

    def test_valid_boolean_values(self) -> None:
        """Test validation of valid boolean values."""
        validator = BooleanTypeValidator()

        for value in ["true", "false", "1", "0"]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"

    def test_invalid_boolean_values(self) -> None:
        """Test validation of invalid boolean values."""
        validator = BooleanTypeValidator()

        for value in ["yes", "no", "2"]:
            result = validator.validate(value)
            assert not result.is_valid, f"{value} should be invalid"


class TestIntegerTypeValidator:
    """Tests for IntegerTypeValidator."""

    def test_valid_integer(self) -> None:
        """Test validation of valid integers."""
        validator = IntegerTypeValidator()

        for value in ["0", "42", "-17", "1000000"]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"

    def test_invalid_integer(self) -> None:
        """Test validation of invalid integers."""
        validator = IntegerTypeValidator()

        for value in ["1.5", "abc", ""]:
            result = validator.validate(value)
            assert not result.is_valid, f"{value} should be invalid"

    def test_min_value(self) -> None:
        """Test min value constraint."""
        validator = IntegerTypeValidator(min_value=0)

        result = validator.validate("0")
        assert result.is_valid

        result = validator.validate("100")
        assert result.is_valid

        result = validator.validate("-1")
        assert not result.is_valid

    def test_max_value(self) -> None:
        """Test max value constraint."""
        validator = IntegerTypeValidator(max_value=100)

        result = validator.validate("100")
        assert result.is_valid

        result = validator.validate("0")
        assert result.is_valid

        result = validator.validate("101")
        assert not result.is_valid

    def test_min_max_range(self) -> None:
        """Test combined min/max range."""
        validator = IntegerTypeValidator(min_value=1, max_value=10)

        result = validator.validate("5")
        assert result.is_valid

        result = validator.validate("0")
        assert not result.is_valid

        result = validator.validate("11")
        assert not result.is_valid


class TestDecimalTypeValidator:
    """Tests for DecimalTypeValidator."""

    def test_valid_decimal(self) -> None:
        """Test validation of valid decimals."""
        validator = DecimalTypeValidator()

        for value in ["0", "1.5", "-3.14", "1000.0", "42"]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"

    def test_invalid_decimal(self) -> None:
        """Test validation of invalid decimals."""
        validator = DecimalTypeValidator()

        for value in ["abc", "", "1.2.3"]:
            result = validator.validate(value)
            assert not result.is_valid, f"{value} should be invalid"

    def test_min_value(self) -> None:
        """Test min value constraint."""
        validator = DecimalTypeValidator(min_value=0.0)

        result = validator.validate("0.5")
        assert result.is_valid

        result = validator.validate("-0.1")
        assert not result.is_valid

    def test_max_value(self) -> None:
        """Test max value constraint."""
        validator = DecimalTypeValidator(max_value=1.0)

        result = validator.validate("0.5")
        assert result.is_valid

        result = validator.validate("1.1")
        assert not result.is_valid


class TestDateTimeTypeValidator:
    """Tests for DateTimeTypeValidator."""

    def test_valid_datetime(self) -> None:
        """Test validation of valid dateTime values."""
        validator = DateTimeTypeValidator()

        for value in [
            "2023-03-25T12:30:45",
            "2020-02-29T23:59:59Z",
            "2020-02-29T23:59:59.123+05:30",
        ]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"

    def test_invalid_datetime(self) -> None:
        """Test validation of invalid dateTime values."""
        validator = DateTimeTypeValidator()

        for value in [
            "2023-02-29T12:00:00",  # invalid day
            "2023-13-01T00:00:00",  # invalid month
            "2023-01-01T24:00:00",  # invalid hour
            "2023-01-01T00:60:00",  # invalid minute
            "2023-01-01 00:00:00",  # missing T separator
            "2023-01-01T00:00",  # missing seconds
            "2023-01-01T00:00:00+14:30",  # invalid timezone minutes for 14
        ]:
            result = validator.validate(value)
            assert not result.is_valid, f"{value} should be invalid"


class TestHexBinaryTypeValidator:
    """Tests for HexBinaryTypeValidator."""

    def test_valid_hex(self) -> None:
        """Test validation of valid hex strings."""
        validator = HexBinaryTypeValidator()

        for value in ["00", "FF", "0123456789ABCDEF", "abcdef"]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"

    def test_invalid_hex(self) -> None:
        """Test validation of invalid hex strings."""
        validator = HexBinaryTypeValidator()

        for value in ["G", "0x00", "123"]:  # Odd length is typically invalid
            if len(value) % 2 != 0:
                result = validator.validate(value)
                assert not result.is_valid, f"{value} (odd length) should be invalid"


class TestNCNameTypeValidator:
    """Tests for NCNameTypeValidator (XML names)."""

    def test_valid_ncname(self) -> None:
        """Test validation of valid NCNames."""
        validator = NCNameTypeValidator()

        for value in ["name", "_name", "Name123", "my-name", "my.name"]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"

    def test_invalid_ncname(self) -> None:
        """Test validation of invalid NCNames."""
        validator = NCNameTypeValidator()

        for value in ["123name", "-name", ":name", "name:space"]:
            result = validator.validate(value)
            assert not result.is_valid, f"{value} should be invalid"


class TestAnyURITypeValidator:
    """Tests for AnyURITypeValidator."""

    def test_valid_uri(self) -> None:
        """Test validation of valid URIs."""
        validator = AnyURITypeValidator()

        for value in [
            "http://example.com",
            "https://example.com/path",
            "file:///path/to/file",
            "../relative/path",
            "/absolute/path",
        ]:
            result = validator.validate(value)
            assert result.is_valid, f"{value} should be valid"
