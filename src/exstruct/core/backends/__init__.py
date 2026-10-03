"""Backend implementations exposed by the core extraction pipeline."""

from __future__ import annotations

from typing import Any

from .base import Backend as Backend

__all__ = [
    "Backend",
    "ComBackend",
    "ComRichBackend",
    "LibreOfficeRichBackend",
    "OoxmlRichBackend",
    "OpenpyxlBackend",
]


def __getattr__(name: str) -> Any:  # noqa: ANN401 - heterogeneous backend exports
    """Resolve the requested backend export without importing other backends."""
    if name in {"ComBackend", "ComRichBackend"}:
        from . import com_backend

        return getattr(com_backend, name)
    if name == "LibreOfficeRichBackend":
        from . import libreoffice_backend

        return libreoffice_backend.LibreOfficeRichBackend
    if name == "OoxmlRichBackend":
        from . import ooxml_backend

        return ooxml_backend.OoxmlRichBackend
    if name == "OpenpyxlBackend":
        from . import openpyxl_backend

        return openpyxl_backend.OpenpyxlBackend
    raise AttributeError(name)
