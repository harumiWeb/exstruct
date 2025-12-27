from __future__ import annotations

import logging

from ..errors import FallbackReason


def log_fallback(logger: logging.Logger, reason: FallbackReason, message: str) -> None:
    """Log a standardized fallback warning.

    Args:
        logger: Logger instance to emit the warning.
        reason: Fallback reason code.
        message: Human-readable detail message.
    """
    logger.warning("[%s] %s", reason.value, message)
