from __future__ import annotations

from pathlib import Path

from exstruct.mcp import extract_runner


def test_apply_conflict_policy_no_conflict(tmp_path: Path) -> None:
    output = tmp_path / "out.json"
    resolved, warning, skipped = extract_runner._apply_conflict_policy(
        output, "overwrite"
    )
    assert resolved == output
    assert warning is None
    assert skipped is False


def test_apply_conflict_policy_skip(tmp_path: Path) -> None:
    output = tmp_path / "out.json"
    output.write_text("data", encoding="utf-8")
    resolved, warning, skipped = extract_runner._apply_conflict_policy(output, "skip")
    assert resolved == output
    assert skipped is True
    assert warning is not None


def test_apply_conflict_policy_rename(tmp_path: Path) -> None:
    output = tmp_path / "out.json"
    output.write_text("data", encoding="utf-8")
    resolved, warning, skipped = extract_runner._apply_conflict_policy(output, "rename")
    assert resolved != output
    assert resolved.name.startswith("out_")
    assert resolved.suffix == ".json"
    assert skipped is False
    assert warning is not None


def test_apply_conflict_policy_overwrite(tmp_path: Path) -> None:
    output = tmp_path / "out.json"
    output.write_text("data", encoding="utf-8")
    resolved, warning, skipped = extract_runner._apply_conflict_policy(
        output, "overwrite"
    )
    assert resolved == output
    assert warning is None
    assert skipped is False
