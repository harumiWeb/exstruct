"""Deferred module access for compatibility and runtime annotation resolution."""

from importlib import import_module
from typing import Any


class LazyModule:
    """Load a dependency only when one of its attributes is actually requested."""

    def __init__(self, name: str) -> None:
        object.__setattr__(self, "_name", name)

    def __getattr__(self, name: str) -> Any:  # noqa: ANN401 - dependency attributes vary
        return getattr(import_module(self._name), name)

    def __setattr__(self, name: str, value: object) -> None:
        """Keep legacy monkeypatches visible to other users of the real module."""
        setattr(import_module(self._name), name, value)

    def __delattr__(self, name: str) -> None:
        delattr(import_module(self._name), name)
