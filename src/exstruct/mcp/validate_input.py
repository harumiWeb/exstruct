from __future__ import annotations

from pathlib import Path

from pydantic import BaseModel, Field

from exstruct.cli.availability import get_com_availability

from .io import PathPolicy

_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


class ValidateInputRequest(BaseModel):
    """Input model for validating an Excel file."""

    xlsx_path: Path


class ValidateInputResult(BaseModel):
    """Output model for input validation."""

    is_readable: bool
    warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)


def validate_input(
    request: ValidateInputRequest, *, policy: PathPolicy | None = None
) -> ValidateInputResult:
    """Validate input Excel file for MCP usage.

    Args:
        request: Validation request payload.
        policy: Optional path policy for access control.

    Returns:
        Validation result with errors and warnings.
    """
    warnings: list[str] = []
    errors: list[str] = []

    try:
        resolved = (
            policy.ensure_allowed(request.xlsx_path)
            if policy
            else request.xlsx_path.resolve()
        )
    except ValueError as exc:
        errors.append(str(exc))
        return ValidateInputResult(is_readable=False, warnings=warnings, errors=errors)

    if not resolved.exists():
        errors.append(f"File not found: {resolved}")
        return ValidateInputResult(is_readable=False, warnings=warnings, errors=errors)

    if not resolved.is_file():
        errors.append(f"Path is not a file: {resolved}")
        return ValidateInputResult(is_readable=False, warnings=warnings, errors=errors)

    if resolved.suffix.lower() not in _ALLOWED_EXTENSIONS:
        errors.append(f"Unsupported file extension: {resolved.suffix}")
        return ValidateInputResult(is_readable=False, warnings=warnings, errors=errors)

    try:
        with resolved.open("rb") as handle:
            handle.read(1)
    except OSError as exc:
        errors.append(f"Failed to read file: {exc}")
        return ValidateInputResult(is_readable=False, warnings=warnings, errors=errors)

    com = get_com_availability()
    if not com.available and com.reason:
        warnings.append(f"COM unavailable: {com.reason}")

    return ValidateInputResult(is_readable=True, warnings=warnings, errors=errors)
