"""Backend implementations exposed by the core extraction pipeline."""

from __future__ import annotations

from .base import Backend
from .com_backend import ComBackend, ComRichBackend
from .libreoffice_backend import LibreOfficeRichBackend
from .ooxml_backend import OoxmlRichBackend
from .openpyxl_backend import OpenpyxlBackend

__all__ = [
    "Backend",
    "ComBackend",
    "ComRichBackend",
    "LibreOfficeRichBackend",
    "OoxmlRichBackend",
    "OpenpyxlBackend",
]
