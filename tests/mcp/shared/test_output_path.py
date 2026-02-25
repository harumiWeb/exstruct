from __future__ import annotations

from pathlib import Path

from exstruct.mcp.shared.output_path import (
    apply_conflict_policy,
    next_available_path,
    normalize_output_name,
)


def test_normalize_output_name_same_stem() -> None:
    input_path = Path("report.xlsx")
    assert (
        normalize_output_name(
            input_path,
            out_name=None,
            default_suffix=".json",
            default_name_builder="same_stem",
        )
        == "report.json"
    )


def test_normalize_output_name_patched() -> None:
    input_path = Path("report.xlsx")
    assert (
        normalize_output_name(
            input_path,
            out_name=None,
            default_suffix=".xlsx",
            default_name_builder="patched",
        )
        == "report_patched.xlsx"
    )


def test_apply_conflict_policy_rename(tmp_path: Path) -> None:
    target = tmp_path / "result.json"
    target.write_text("x", encoding="utf-8")
    resolved, warning, skipped = apply_conflict_policy(target, "rename")
    assert resolved.name == "result_1.json"
    assert warning == "Output exists; renamed to: result_1.json"
    assert skipped is False


def test_next_available_path_no_conflict(tmp_path: Path) -> None:
    target = tmp_path / "result.json"
    assert next_available_path(target) == target
