from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp.io import PathPolicy


def test_path_policy_allows_within_root(tmp_path: Path) -> None:
    policy = PathPolicy(root=tmp_path)
    target = tmp_path / "data" / "file.txt"
    allowed = policy.ensure_allowed(target)
    assert allowed == target.resolve()


def test_path_policy_denies_outside_root(tmp_path: Path) -> None:
    policy = PathPolicy(root=tmp_path)
    outside = tmp_path.parent / "outside.txt"
    with pytest.raises(ValueError):
        policy.ensure_allowed(outside)


def test_path_policy_denies_glob(tmp_path: Path) -> None:
    policy = PathPolicy(root=tmp_path, deny_globs=["**/*.secret"])
    denied = tmp_path / "nested" / "token.secret"
    denied.parent.mkdir(parents=True, exist_ok=True)
    denied.write_text("x", encoding="utf-8")
    with pytest.raises(ValueError):
        policy.ensure_allowed(denied)
