from __future__ import annotations

from pathlib import Path

from exstruct.mcp.shared.output_path import (
    apply_conflict_policy,
    next_available_directory,
    next_available_path,
    normalize_output_name,
    resolve_image_output_dir,
)


def test_normalize_output_name_same_stem() -> None:
    """Use input stem with default suffix in same-stem mode."""
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
    """Append `_patched` once when patched naming is requested."""
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


def test_normalize_output_name_patched_does_not_chain_suffix() -> None:
    """Avoid repeatedly appending `_patched`."""
    input_path = Path("report_patched.xlsx")
    assert (
        normalize_output_name(
            input_path,
            out_name=None,
            default_suffix=".xlsx",
            default_name_builder="patched",
        )
        == "report_patched.xlsx"
    )


def test_normalize_output_name_patched_case_insensitive() -> None:
    """Treat `_patched` suffix detection as case-insensitive."""
    input_path = Path("report_Patched.xlsx")
    assert (
        normalize_output_name(
            input_path,
            out_name=None,
            default_suffix=".xlsx",
            default_name_builder="patched",
        )
        == "report_Patched.xlsx"
    )


def test_apply_conflict_policy_rename(tmp_path: Path) -> None:
    """Rename output when conflict policy is `rename`."""
    target = tmp_path / "result.json"
    target.write_text("x", encoding="utf-8")
    resolved, warning, skipped = apply_conflict_policy(target, "rename")
    assert resolved.name == "result_1.json"
    assert warning == "Output exists; renamed to: result_1.json"
    assert skipped is False


def test_next_available_path_no_conflict(tmp_path: Path) -> None:
    """Return unchanged path when file path is unused."""
    target = tmp_path / "result.json"
    assert next_available_path(target) == target


def test_resolve_image_output_dir_defaults_to_stem_images(tmp_path: Path) -> None:
    """Reserve `<stem>_images` directory by default."""
    input_path = tmp_path / "book.xlsx"
    resolved = resolve_image_output_dir(input_path, out_dir=None, policy=None)
    assert resolved == (tmp_path / "book_images").resolve()
    assert resolved.exists()
    assert resolved.is_dir()


def test_resolve_image_output_dir_appends_suffix_on_conflict(tmp_path: Path) -> None:
    """Reserve next suffixed directory when default name already exists."""
    input_path = tmp_path / "book.xlsx"
    (tmp_path / "book_images").mkdir(parents=True, exist_ok=True)
    resolved = resolve_image_output_dir(input_path, out_dir=None, policy=None)
    assert resolved == (tmp_path / "book_images_1").resolve()
    assert resolved.exists()
    assert resolved.is_dir()


def test_resolve_image_output_dir_uses_out_dir_when_specified(tmp_path: Path) -> None:
    """Respect explicit output directory without auto-reserving a new one."""
    input_path = tmp_path / "book.xlsx"
    explicit = tmp_path / "custom_images"
    resolved = resolve_image_output_dir(input_path, out_dir=explicit, policy=None)
    assert resolved == explicit.resolve()


def test_next_available_directory_reserves_unique_directories(tmp_path: Path) -> None:
    """Reserve unique directory paths atomically on repeated calls."""
    base = tmp_path / "book_images"
    first = next_available_directory(base, policy=None)
    second = next_available_directory(base, policy=None)
    assert first == base.resolve()
    assert second == (tmp_path / "book_images_1").resolve()
    assert first.exists()
    assert second.exists()
