import logging

import pytest

from exstruct.core.logging_utils import log_fallback
from exstruct.errors import FallbackReason


def test_log_fallback_includes_reason(caplog: pytest.LogCaptureFixture) -> None:
    logger = logging.getLogger("exstruct.test")
    with caplog.at_level(logging.WARNING):
        log_fallback(logger, FallbackReason.LIGHT_MODE, "fallback")

    assert any("[light_mode] fallback" in record.message for record in caplog.records)
