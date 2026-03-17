from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp.io import PathPolicy
from exstruct.mcp.patch import runtime as legacy_runtime


def test_legacy_runtime_resolve_input_path_accepts_policy(tmp_path: Path) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_text("dummy", encoding="utf-8")

    resolved = legacy_runtime.resolve_input_path(
        input_path, policy=PathPolicy(root=tmp_path)
    )

    assert resolved == input_path.resolve()


def test_legacy_runtime_resolve_output_path_accepts_policy(tmp_path: Path) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_text("dummy", encoding="utf-8")

    with pytest.raises(ValueError, match="Path is outside root"):
        legacy_runtime.resolve_output_path(
            input_path,
            out_dir=tmp_path.parent,
            out_name="patched.xlsx",
            policy=PathPolicy(root=tmp_path),
        )


def test_legacy_runtime_resolve_make_output_path_accepts_policy(tmp_path: Path) -> None:
    resolved = legacy_runtime.resolve_make_output_path(
        Path("nested") / "book.xlsx",
        policy=PathPolicy(root=tmp_path),
    )

    assert resolved == (tmp_path / "nested" / "book.xlsx").resolve()
