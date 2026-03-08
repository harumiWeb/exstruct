"""Tests for LibreOffice-specific request validation constraints."""

from pathlib import Path

import pytest

from exstruct import ConfigError
from exstruct.constraints import (
    validate_libreoffice_extraction_request,
    validate_libreoffice_process_request,
)


def test_validate_libreoffice_extraction_request_rejects_auto_page_breaks() -> None:
    """Verify that LibreOffice extraction rejects auto page-break requests."""

    with pytest.raises(ConfigError, match="does not support auto page-break export"):
        validate_libreoffice_extraction_request(
            Path("book.xlsx"),
            mode="libreoffice",
            include_auto_page_breaks=True,
        )


def test_validate_libreoffice_extraction_request_rejects_xls() -> None:
    """Verify that LibreOffice extraction rejects legacy .xls workbooks."""

    with pytest.raises(ValueError, match="not supported in libreoffice mode"):
        validate_libreoffice_extraction_request(
            Path("book.xls"),
            mode="libreoffice",
        )


def test_validate_libreoffice_extraction_request_allows_non_libreoffice_passthrough() -> (
    None
):
    """Verify that non-LibreOffice extraction keeps auto page-break requests."""

    normalized = validate_libreoffice_extraction_request(
        Path("book.xlsx"),
        mode="standard",
        include_auto_page_breaks=True,
    )
    assert normalized == Path("book.xlsx")


def test_validate_libreoffice_process_request_rejects_auto_page_breaks_only() -> None:
    """Verify that validate LibreOffice process request rejects auto page breaks only."""

    with pytest.raises(ConfigError, match="does not support auto page-break export"):
        validate_libreoffice_process_request(
            Path("book.xlsx"),
            mode="libreoffice",
            include_auto_page_breaks=True,
        )


def test_validate_libreoffice_process_request_prefers_combined_error() -> None:
    """Verify that validate LibreOffice process request prefers combined error."""

    with pytest.raises(
        ConfigError,
        match="does not support PDF/PNG rendering or auto page-break export",
    ):
        validate_libreoffice_process_request(
            Path("book.xlsx"),
            mode="libreoffice",
            include_auto_page_breaks=True,
            pdf=True,
        )


def test_validate_libreoffice_process_request_rejects_image_only() -> None:
    """Verify that LibreOffice process requests reject image rendering."""

    with pytest.raises(ConfigError, match="does not support PDF/PNG rendering"):
        validate_libreoffice_process_request(
            Path("book.xlsx"),
            mode="libreoffice",
            image=True,
        )


def test_validate_libreoffice_process_request_allows_standard_passthrough() -> None:
    """Verify that non-LibreOffice process requests keep rendering flags."""

    normalized = validate_libreoffice_process_request(
        Path("book.xlsx"),
        mode="standard",
        include_auto_page_breaks=True,
        image=True,
    )
    assert normalized == Path("book.xlsx")


def test_validate_libreoffice_process_request_allows_string_path_without_flags() -> (
    None
):
    """Verify that valid LibreOffice process requests normalize string paths."""

    normalized = validate_libreoffice_process_request(
        "book.xlsx",
        mode="libreoffice",
    )
    assert normalized == Path("book.xlsx")
