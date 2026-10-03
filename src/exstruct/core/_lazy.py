"""Deferred module access for compatibility and runtime annotation resolution."""

from typing import Any, Literal


class LazyModule:
    """Load xlwings only when one of its attributes is actually requested."""

    def __init__(self, name: Literal["xlwings"]) -> None:
        if name != "xlwings":
            raise ValueError("Only the xlwings compatibility module is supported")

    def __getattr__(self, name: str) -> Any:  # noqa: ANN401 - dependency attributes vary
        import xlwings

        return getattr(xlwings, name)

    def __setattr__(self, name: str, value: object) -> None:
        """Keep legacy monkeypatches visible to other users of the real module."""
        import xlwings

        setattr(xlwings, name, value)

    def __delattr__(self, name: str) -> None:
        import xlwings

        delattr(xlwings, name)
