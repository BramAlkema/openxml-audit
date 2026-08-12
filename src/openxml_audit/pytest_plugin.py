"""pytest plugin for openxml-audit.

Auto-registers fixtures when openxml-audit is installed. No conftest.py
wiring required — just ``pip install openxml-audit`` and use the fixtures.

Fixtures
--------
openxml_validator
    Pre-configured ``OpenXmlValidator`` instance. Respects ``--openxml-format``.
assert_valid_pptx / assert_valid_docx / assert_valid_xlsx
    Call with a file path to assert the file is valid OOXML.
assert_valid_odf
    Call with a file path to assert the file is valid ODF.

CLI options
-----------
--openxml-format    Office version to validate against (default: Office2019).
--openxml-max-errors  Maximum errors to collect (default: 100).
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest

from openxml_audit.errors import FileFormat
from openxml_audit.validator import OpenXmlValidator

_FORMAT_MAP = {
    "Office2007": FileFormat.OFFICE_2007,
    "Office2010": FileFormat.OFFICE_2010,
    "Office2013": FileFormat.OFFICE_2013,
    "Office2016": FileFormat.OFFICE_2016,
    "Office2019": FileFormat.OFFICE_2019,
    "Microsoft365": FileFormat.MICROSOFT_365,
}


def pytest_addoption(parser: pytest.Parser) -> None:
    group = parser.getgroup("openxml-audit", "OpenXML / ODF file validation")
    group.addoption(
        "--openxml-format",
        default="Office2019",
        choices=list(_FORMAT_MAP.keys()),
        help="Office version to validate against (default: Office2019)",
    )
    group.addoption(
        "--openxml-max-errors",
        default=100,
        type=int,
        help="Maximum validation errors to collect (default: 100)",
    )


@pytest.fixture(scope="session")
def openxml_validator(request: pytest.FixtureRequest) -> OpenXmlValidator:
    """Pre-configured OpenXmlValidator instance."""
    fmt = _FORMAT_MAP[request.config.getoption("--openxml-format")]
    max_errors = request.config.getoption("--openxml-max-errors")
    return OpenXmlValidator(file_format=fmt, max_errors=max_errors)


def _make_assert_valid(validator: OpenXmlValidator) -> Callable[[Path | str], None]:
    def _assert(path: Path | str) -> None:
        result = validator.validate(path)
        if not result.is_valid:
            lines = [f"  - {e.description}" for e in result.errors[:10]]
            if len(result.errors) > 10:
                lines.append(f"  ... (+{len(result.errors) - 10} more)")
            pytest.fail(f"Validation failed ({result.error_count} errors):\n" + "\n".join(lines))

    return _assert


@pytest.fixture(scope="session")
def assert_valid_pptx(openxml_validator: OpenXmlValidator) -> Callable[[Path | str], None]:
    """Assert a PPTX file is valid. Usage: ``assert_valid_pptx("out.pptx")``."""
    return _make_assert_valid(openxml_validator)


@pytest.fixture(scope="session")
def assert_valid_docx(openxml_validator: OpenXmlValidator) -> Callable[[Path | str], None]:
    """Assert a DOCX file is valid. Usage: ``assert_valid_docx("out.docx")``."""
    return _make_assert_valid(openxml_validator)


@pytest.fixture(scope="session")
def assert_valid_xlsx(openxml_validator: OpenXmlValidator) -> Callable[[Path | str], None]:
    """Assert an XLSX file is valid. Usage: ``assert_valid_xlsx("out.xlsx")``."""
    return _make_assert_valid(openxml_validator)


@pytest.fixture(scope="session")
def assert_valid_odf(request: pytest.FixtureRequest) -> Callable[[Path | str], None]:
    """Assert an ODF file is valid. Usage: ``assert_valid_odf("doc.odt")``."""
    from openxml_audit.odf import OdfValidator

    validator = OdfValidator()

    def _assert(path: Path | str) -> None:
        result = validator.validate(path)
        if not result.is_valid:
            lines = [f"  - {e.description}" for e in result.errors[:10]]
            if len(result.errors) > 10:
                lines.append(f"  ... (+{len(result.errors) - 10} more)")
            message = f"ODF validation failed ({result.error_count} errors):\n" + "\n".join(lines)
            pytest.fail(message)

    return _assert
