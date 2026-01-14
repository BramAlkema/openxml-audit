"""XSD type validators for attribute and element content validation."""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from enum import Enum
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from openxml_audit.context import ValidationContext


class XsdBuiltinType(Enum):
    """XSD built-in types."""

    STRING = "string"
    NORMALIZED_STRING = "normalizedString"
    TOKEN = "token"
    BOOLEAN = "boolean"
    INTEGER = "integer"
    POSITIVE_INTEGER = "positiveInteger"
    NON_NEGATIVE_INTEGER = "nonNegativeInteger"
    NEGATIVE_INTEGER = "negativeInteger"
    NON_POSITIVE_INTEGER = "nonPositiveInteger"
    LONG = "long"
    INT = "int"
    SHORT = "short"
    BYTE = "byte"
    UNSIGNED_LONG = "unsignedLong"
    UNSIGNED_INT = "unsignedInt"
    UNSIGNED_SHORT = "unsignedShort"
    UNSIGNED_BYTE = "unsignedByte"
    DECIMAL = "decimal"
    FLOAT = "float"
    DOUBLE = "double"
    DATETIME = "dateTime"
    DATE = "date"
    TIME = "time"
    DURATION = "duration"
    HEX_BINARY = "hexBinary"
    BASE64_BINARY = "base64Binary"
    ANY_URI = "anyURI"
    QNAME = "QName"
    ID = "ID"
    IDREF = "IDREF"
    NCNAME = "NCName"
    NMTOKEN = "NMTOKEN"


@dataclass
class TypeValidationResult:
    """Result of type validation."""

    is_valid: bool
    error_message: str | None = None
    parsed_value: Any = None


class XsdTypeValidator(ABC):
    """Base class for XSD type validators."""

    @abstractmethod
    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        """Validate a string value against this type.

        Args:
            value: The string value to validate.
            context: Optional validation context.

        Returns:
            TypeValidationResult with validation status.
        """
        pass


class StringTypeValidator(XsdTypeValidator):
    """Validates string types with optional constraints."""

    def __init__(
        self,
        min_length: int | None = None,
        max_length: int | None = None,
        pattern: str | None = None,
        enumeration: list[str] | None = None,
    ):
        self.min_length = min_length
        self.max_length = max_length
        self.pattern = re.compile(pattern) if pattern else None
        self.enumeration = set(enumeration) if enumeration else None

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        # Check length constraints
        if self.min_length is not None and len(value) < self.min_length:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"String length {len(value)} is less than minimum {self.min_length}",
            )

        if self.max_length is not None and len(value) > self.max_length:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"String length {len(value)} exceeds maximum {self.max_length}",
            )

        # Check pattern
        if self.pattern is not None and not self.pattern.match(value):
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Value '{value}' does not match required pattern",
            )

        # Check enumeration
        if self.enumeration is not None and value not in self.enumeration:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Value '{value}' is not in allowed values: {sorted(self.enumeration)}",
            )

        return TypeValidationResult(is_valid=True, parsed_value=value)


class BooleanTypeValidator(XsdTypeValidator):
    """Validates boolean values."""

    VALID_VALUES = {"true", "false", "1", "0"}

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        if value.lower() not in self.VALID_VALUES:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid boolean value: '{value}'. Expected true, false, 1, or 0",
            )

        parsed = value.lower() in ("true", "1")
        return TypeValidationResult(is_valid=True, parsed_value=parsed)


class IntegerTypeValidator(XsdTypeValidator):
    """Validates integer values with optional min/max constraints."""

    def __init__(
        self,
        min_value: int | None = None,
        max_value: int | None = None,
        min_inclusive: bool = True,
        max_inclusive: bool = True,
    ):
        self.min_value = min_value
        self.max_value = max_value
        self.min_inclusive = min_inclusive
        self.max_inclusive = max_inclusive

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        try:
            parsed = int(value)
        except ValueError:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid integer value: '{value}'",
            )

        # Check min constraint
        if self.min_value is not None:
            if self.min_inclusive and parsed < self.min_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} is less than minimum {self.min_value}",
                )
            if not self.min_inclusive and parsed <= self.min_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} must be greater than {self.min_value}",
                )

        # Check max constraint
        if self.max_value is not None:
            if self.max_inclusive and parsed > self.max_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} exceeds maximum {self.max_value}",
                )
            if not self.max_inclusive and parsed >= self.max_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} must be less than {self.max_value}",
                )

        return TypeValidationResult(is_valid=True, parsed_value=parsed)


class DecimalTypeValidator(XsdTypeValidator):
    """Validates decimal values."""

    def __init__(
        self,
        min_value: Decimal | float | None = None,
        max_value: Decimal | float | None = None,
        min_inclusive: bool = True,
        max_inclusive: bool = True,
        total_digits: int | None = None,
        fraction_digits: int | None = None,
    ):
        self.min_value = Decimal(str(min_value)) if min_value is not None else None
        self.max_value = Decimal(str(max_value)) if max_value is not None else None
        self.min_inclusive = min_inclusive
        self.max_inclusive = max_inclusive
        self.total_digits = total_digits
        self.fraction_digits = fraction_digits

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        try:
            parsed = Decimal(value)
        except InvalidOperation:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid decimal value: '{value}'",
            )

        # Check min constraint
        if self.min_value is not None:
            if self.min_inclusive and parsed < self.min_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} is less than minimum {self.min_value}",
                )
            if not self.min_inclusive and parsed <= self.min_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} must be greater than {self.min_value}",
                )

        # Check max constraint
        if self.max_value is not None:
            if self.max_inclusive and parsed > self.max_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} exceeds maximum {self.max_value}",
                )
            if not self.max_inclusive and parsed >= self.max_value:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Value {parsed} must be less than {self.max_value}",
                )

        return TypeValidationResult(is_valid=True, parsed_value=parsed)


class DateTimeTypeValidator(XsdTypeValidator):
    """Validates XSD dateTime values."""

    DATETIME_PATTERN = re.compile(
        r"^(?P<year>-?\d{4,})-(?P<month>\d{2})-(?P<day>\d{2})"
        r"T(?P<hour>\d{2}):(?P<minute>\d{2}):(?P<second>\d{2})"
        r"(?P<fraction>\.\d+)?(?P<tz>Z|[+-]\d{2}:\d{2})?$"
    )

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        match = self.DATETIME_PATTERN.match(value)
        if not match:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid dateTime value: '{value}'",
            )

        year = int(match.group("year"))
        month = int(match.group("month"))
        day = int(match.group("day"))
        hour = int(match.group("hour"))
        minute = int(match.group("minute"))
        second = int(match.group("second"))
        fraction = match.group("fraction")
        tz = match.group("tz")

        if not 1 <= month <= 12:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid month in dateTime value: '{value}'",
            )

        max_day = self._days_in_month(year, month)
        if not 1 <= day <= max_day:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid day in dateTime value: '{value}'",
            )

        if not 0 <= hour <= 23:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid hour in dateTime value: '{value}'",
            )

        if not 0 <= minute <= 59:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid minute in dateTime value: '{value}'",
            )

        if not 0 <= second <= 60:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid second in dateTime value: '{value}'",
            )

        tzinfo = None
        if tz:
            if tz == "Z":
                tzinfo = timezone.utc
            else:
                sign = 1 if tz[0] == "+" else -1
                tz_hour = int(tz[1:3])
                tz_minute = int(tz[4:6])
                if not 0 <= tz_hour <= 14 or not 0 <= tz_minute <= 59:
                    return TypeValidationResult(
                        is_valid=False,
                        error_message=f"Invalid timezone in dateTime value: '{value}'",
                    )
                if tz_hour == 14 and tz_minute != 0:
                    return TypeValidationResult(
                        is_valid=False,
                        error_message=f"Invalid timezone in dateTime value: '{value}'",
                    )
                offset = timedelta(hours=tz_hour, minutes=tz_minute) * sign
                tzinfo = timezone(offset)

        parsed = None
        if 1 <= year <= 9999 and second <= 59:
            microsecond = 0
            if fraction:
                fraction_digits = fraction[1:]
                microsecond = int((fraction_digits + "000000")[:6])
            try:
                parsed = datetime(
                    year,
                    month,
                    day,
                    hour,
                    minute,
                    second,
                    microsecond,
                    tzinfo=tzinfo,
                )
            except ValueError:
                return TypeValidationResult(
                    is_valid=False,
                    error_message=f"Invalid dateTime value: '{value}'",
                )

        return TypeValidationResult(is_valid=True, parsed_value=parsed or value)

    def _days_in_month(self, year: int, month: int) -> int:
        if month == 2:
            return 29 if self._is_leap_year(year) else 28
        if month in (4, 6, 9, 11):
            return 30
        return 31

    def _is_leap_year(self, year: int) -> bool:
        if year % 400 == 0:
            return True
        if year % 100 == 0:
            return False
        return year % 4 == 0


class HexBinaryTypeValidator(XsdTypeValidator):
    """Validates hexBinary values."""

    HEX_PATTERN = re.compile(r"^[0-9A-Fa-f]*$")

    def __init__(self, length: int | None = None):
        self.length = length

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        if not self.HEX_PATTERN.match(value):
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid hexBinary value: '{value}'",
            )

        if len(value) % 2 != 0:
            return TypeValidationResult(
                is_valid=False,
                error_message="hexBinary value must have even number of characters",
            )

        if self.length is not None and len(value) // 2 != self.length:
            return TypeValidationResult(
                is_valid=False,
                error_message=f"hexBinary length {len(value) // 2} does not match required {self.length}",
            )

        return TypeValidationResult(is_valid=True, parsed_value=bytes.fromhex(value))


class NCNameTypeValidator(XsdTypeValidator):
    """Validates NCName (non-colonized name) values."""

    # NCName pattern: starts with letter or underscore, followed by letters, digits, hyphens, underscores, periods
    NCNAME_PATTERN = re.compile(r"^[a-zA-Z_][\w.\-]*$")

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        if not value:
            return TypeValidationResult(
                is_valid=False,
                error_message="NCName cannot be empty",
            )

        if not self.NCNAME_PATTERN.match(value):
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid NCName: '{value}'",
            )

        return TypeValidationResult(is_valid=True, parsed_value=value)


class AnyURITypeValidator(XsdTypeValidator):
    """Validates anyURI values."""

    def validate(self, value: str, context: ValidationContext | None = None) -> TypeValidationResult:
        # Basic URI validation - allow most strings that could be URIs
        # A more complete implementation would parse according to RFC 3986
        if not value:
            return TypeValidationResult(is_valid=True, parsed_value=value)

        # Check for invalid characters
        invalid_chars = set('<>"{}|\\^`')
        if any(c in value for c in invalid_chars):
            return TypeValidationResult(
                is_valid=False,
                error_message=f"Invalid URI: contains invalid characters",
            )

        return TypeValidationResult(is_valid=True, parsed_value=value)


# Pre-built validators for common XSD types
BUILTIN_VALIDATORS: dict[XsdBuiltinType, XsdTypeValidator] = {
    XsdBuiltinType.STRING: StringTypeValidator(),
    XsdBuiltinType.BOOLEAN: BooleanTypeValidator(),
    XsdBuiltinType.INTEGER: IntegerTypeValidator(),
    XsdBuiltinType.POSITIVE_INTEGER: IntegerTypeValidator(min_value=1),
    XsdBuiltinType.NON_NEGATIVE_INTEGER: IntegerTypeValidator(min_value=0),
    XsdBuiltinType.NEGATIVE_INTEGER: IntegerTypeValidator(max_value=-1),
    XsdBuiltinType.NON_POSITIVE_INTEGER: IntegerTypeValidator(max_value=0),
    XsdBuiltinType.LONG: IntegerTypeValidator(min_value=-9223372036854775808, max_value=9223372036854775807),
    XsdBuiltinType.INT: IntegerTypeValidator(min_value=-2147483648, max_value=2147483647),
    XsdBuiltinType.SHORT: IntegerTypeValidator(min_value=-32768, max_value=32767),
    XsdBuiltinType.BYTE: IntegerTypeValidator(min_value=-128, max_value=127),
    XsdBuiltinType.UNSIGNED_LONG: IntegerTypeValidator(min_value=0, max_value=18446744073709551615),
    XsdBuiltinType.UNSIGNED_INT: IntegerTypeValidator(min_value=0, max_value=4294967295),
    XsdBuiltinType.UNSIGNED_SHORT: IntegerTypeValidator(min_value=0, max_value=65535),
    XsdBuiltinType.UNSIGNED_BYTE: IntegerTypeValidator(min_value=0, max_value=255),
    XsdBuiltinType.DECIMAL: DecimalTypeValidator(),
    XsdBuiltinType.DATETIME: DateTimeTypeValidator(),
    XsdBuiltinType.HEX_BINARY: HexBinaryTypeValidator(),
    XsdBuiltinType.NCNAME: NCNameTypeValidator(),
    XsdBuiltinType.ID: NCNameTypeValidator(),
    XsdBuiltinType.IDREF: NCNameTypeValidator(),
    XsdBuiltinType.ANY_URI: AnyURITypeValidator(),
}


def get_type_validator(type_name: str | XsdBuiltinType) -> XsdTypeValidator | None:
    """Get a validator for an XSD built-in type.

    Args:
        type_name: The type name or XsdBuiltinType enum.

    Returns:
        The appropriate validator, or None if not found.
    """
    if isinstance(type_name, str):
        try:
            type_enum = XsdBuiltinType(type_name)
        except ValueError:
            return None
    else:
        type_enum = type_name

    return BUILTIN_VALIDATORS.get(type_enum)
