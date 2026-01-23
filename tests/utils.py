from collections.abc import Callable, Iterable, Sequence
from typing import Literal, TypeVar, cast

import pytest
from typing_extensions import ParamSpec

P = ParamSpec("P")
R = TypeVar("R")


def parametrize(
    argnames: str | Sequence[str],
    argvalues: Iterable[object],
    *,
    indirect: bool | Sequence[str] = False,
    ids: Iterable[str | float | int | bool | None]
    | Callable[[object], object | None]
    | None = None,
    scope: Literal["session", "package", "module", "class", "function"] | None = None,
) -> Callable[[Callable[P, R]], Callable[P, R]]:
    """
    Return a decorator that parametrizes a test callable with the given argument names and values.
    
    Parameters:
        argnames: One or more parameter names (single string or sequence of strings) to inject into the test callable.
        argvalues: An iterable of values or value-tuples to use for each generated test case.
        indirect: If True or a sequence of names, treat corresponding parameters as fixtures and resolve them indirectly.
        ids: Optional iterable of case identifiers or a callable that produces an identifier for each value.
        scope: Optional fixture scope to apply when parameters are used as fixtures ("session", "package", "module", "class", or "function").
    
    Returns:
        decorator: A decorator that applies the specified parametrization to a callable while preserving its signature.
    """
    return cast(
        Callable[[Callable[P, R]], Callable[P, R]],
        pytest.mark.parametrize(
            argnames,
            argvalues,
            indirect=indirect,
            ids=ids,
            scope=scope,
        ),
    )