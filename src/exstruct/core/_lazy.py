"""Deferred module access for compatibility and runtime annotation resolution."""

from importlib import import_module
from typing import Any


class LazyModule:
    """Load a dependency only when one of its attributes is actually requested."""

    def __init__(self, name: str) -> None:
        self._name = name

    def __getattr__(self, name: str) -> Any:  # noqa: ANN401 - dependency attributes vary
        return getattr(import_module(self._name), name)
