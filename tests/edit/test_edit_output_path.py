from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.edit.output_path import next_available_directory, next_available_path
from exstruct.mcp.io import PathPolicy


def test_next_available_path_reserves_file_atomically(tmp_path: Path) -> None:
    target = tmp_path / "result.json"

    reserved = next_available_path(target)

    assert reserved == target.resolve()
    assert reserved.exists()
    assert reserved.is_file()


def test_next_available_directory_checks_policy_before_reserving(
    tmp_path: Path,
) -> None:
    root = tmp_path / "root"
    root.mkdir()
    outside = tmp_path / "outside" / "images"

    with pytest.raises(ValueError, match="Path is outside root"):
        next_available_directory(outside, policy=PathPolicy(root=root))

    assert not outside.exists()
