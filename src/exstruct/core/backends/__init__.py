"""Backend implementations exposed by the core extraction pipeline."""

from __future__ import annotations

from importlib import import_module
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
    modules = {
        "ComBackend": "com_backend",
        "ComRichBackend": "com_backend",
        "LibreOfficeRichBackend": "libreoffice_backend",
        "OoxmlRichBackend": "ooxml_backend",
        "OpenpyxlBackend": "openpyxl_backend",
    }
    if name not in modules:
        raise AttributeError(name)
    return getattr(import_module(f"{__name__}.{modules[name]}"), name)
