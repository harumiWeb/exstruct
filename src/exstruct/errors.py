from __future__ import annotations

"""Project-specific exception hierarchy for ExStruct."""


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
