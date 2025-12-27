"""Project-specific exception hierarchy for ExStruct."""

from __future__ import annotations

from enum import Enum


class ExstructError(Exception):
    """Base exception for ExStruct."""


class ConfigError(ExstructError):
    """Raised when user-provided configuration or parameters are invalid."""


class ExtractionError(ExstructError):
    """Raised when workbook extraction fails."""


class SerializationError(ExstructError):
    """Raised when serialization fails or an unsupported format is requested."""


class MissingDependencyError(ExstructError):
    """Raised when an optional dependency required for the requested operation is missing."""


class RenderError(ExstructError):
    """Raised when rendering (PDF/PNG) fails."""


class OutputError(ExstructError):
    """Raised when writing outputs to disk or streams fails."""


class PrintAreaError(ExstructError, ValueError):
    """Raised when print-area specific processing fails (also a ValueError for compatibility)."""


class FallbackReason(str, Enum):
    """Reason codes for extraction fallbacks."""

    LIGHT_MODE = "light_mode"
    SKIP_COM_TESTS = "skip_com_tests"
    COM_UNAVAILABLE = "com_unavailable"
    COM_PIPELINE_FAILED = "com_pipeline_failed"
