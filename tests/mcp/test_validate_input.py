from __future__ import annotations

import importlib
from pathlib import Path

import pytest

from exstruct.cli.availability import ComAvailability
from exstruct.mcp.io import PathPolicy
from exstruct.mcp.validate_input import ValidateInputRequest, validate_input


def test_validate_input_missing_file(tmp_path: Path) -> None:
    request = ValidateInputRequest(xlsx_path=tmp_path / "missing.xlsx")
    result = validate_input(request)
    assert result.is_readable is False
    assert result.errors


def test_validate_input_invalid_extension(tmp_path: Path) -> None:
    path = tmp_path / "input.txt"
    path.write_text("x", encoding="utf-8")
    request = ValidateInputRequest(xlsx_path=path)
    result = validate_input(request)
    assert result.is_readable is False
    assert "Unsupported file extension" in result.errors[0]


def test_validate_input_policy_denied(tmp_path: Path) -> None:
    policy = PathPolicy(root=tmp_path)
    outside = tmp_path.parent / "outside.xlsx"
    request = ValidateInputRequest(xlsx_path=outside)
    result = validate_input(request, policy=policy)
    assert result.is_readable is False
    assert result.errors


def test_validate_input_warns_on_com(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "input.xlsx"
    path.write_bytes(b"test")
    validate_input_module = importlib.import_module("exstruct.mcp.validate_input")
    monkeypatch.setattr(
        validate_input_module,
        "get_com_availability",
        lambda: ComAvailability(available=False, reason="No COM"),
    )
    request = ValidateInputRequest(xlsx_path=path)
    result = validate_input(request)
    assert result.is_readable is True
    assert any("COM unavailable" in warning for warning in result.warnings)


def test_validate_input_rejects_directory(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    path.mkdir()
    request = ValidateInputRequest(xlsx_path=path)
    result = validate_input(request)
    assert result.is_readable is False
    assert any("Path is not a file" in error for error in result.errors)


def test_validate_input_handles_read_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "input.xlsx"
    path.write_text("x", encoding="utf-8")

    def _raise(*_args: object, **_kwargs: object) -> object:
        raise OSError("boom")

    monkeypatch.setattr(Path, "open", _raise)
    request = ValidateInputRequest(xlsx_path=path)
    result = validate_input(request)
    assert result.is_readable is False
    assert any("Failed to read file" in error for error in result.errors)
