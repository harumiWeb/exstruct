from __future__ import annotations

from pathlib import Path

from pydantic import BaseModel, Field, field_validator, model_validator

from exstruct import render

from .io import PathPolicy
from .shared.a1 import resolve_sheet_and_range
from .shared.output_path import resolve_image_output_dir


class CaptureSheetImagesRequest(BaseModel):
    """Input model for MCP sheet image capture."""

    xlsx_path: Path
    out_dir: Path | None = None
    dpi: int = Field(default=144, ge=1)
    sheet: str | None = None
    range: str | None = None  # noqa: A003

    @field_validator("sheet")
    @classmethod
    def _validate_sheet(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not candidate:
            raise ValueError("sheet must not be empty when provided.")
        return candidate

    @field_validator("range")
    @classmethod
    def _validate_range(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not candidate:
            raise ValueError("range must not be empty when provided.")
        return candidate

    @model_validator(mode="after")
    def _validate_sheet_range_consistency(self) -> CaptureSheetImagesRequest:
        selection = resolve_sheet_and_range(self.sheet, self.range)
        self.sheet = selection.sheet
        self.range = selection.range_ref
        return self


class CaptureSheetImagesResult(BaseModel):
    """Output model for MCP sheet image capture."""

    out_dir: str
    image_paths: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


def run_capture_sheet_images(
    request: CaptureSheetImagesRequest,
    *,
    policy: PathPolicy | None = None,
) -> CaptureSheetImagesResult:
    """Capture worksheet images using Excel COM rendering pipeline.

    Args:
        request: Capture request payload.
        policy: Optional path access policy.

    Returns:
        Capture result with resolved output directory and written images.
    """
    resolved_input = _resolve_input_path(request.xlsx_path, policy=policy)
    resolved_out_dir = resolve_image_output_dir(
        resolved_input,
        out_dir=request.out_dir,
        policy=policy,
    )
    _ensure_com_available()
    resolved_out_dir.mkdir(parents=True, exist_ok=True)
    written_paths = render.export_sheet_images(
        resolved_input,
        resolved_out_dir,
        dpi=request.dpi,
        sheet=request.sheet,
        a1_range=request.range,
    )
    return CaptureSheetImagesResult(
        out_dir=str(resolved_out_dir),
        image_paths=[str(path) for path in written_paths],
        warnings=[],
    )


def _resolve_input_path(path: Path, *, policy: PathPolicy | None) -> Path:
    """Resolve and validate input workbook path."""
    resolved = policy.ensure_allowed(path) if policy else path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Input file not found: {resolved}")
    if not resolved.is_file():
        raise ValueError(f"Input path is not a file: {resolved}")
    return resolved


def _ensure_com_available() -> None:
    """Validate Excel COM availability and raise ValueError when unavailable."""
    app: object | None = None
    try:
        app = render._require_excel_app()
    except Exception as exc:  # pragma: no cover - delegated by render internals
        raise ValueError(
            "Excel (COM) is not available. Rendering (PDF/image) requires a desktop Excel installation."
        ) from exc
    finally:
        if app is not None:
            quit_method = getattr(app, "quit", None)
            if callable(quit_method):
                quit_method()
