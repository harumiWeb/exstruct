from __future__ import annotations

from pathlib import Path
from typing import Literal

from exstruct.mcp.io import PathPolicy

OnConflictPolicy = Literal["overwrite", "skip", "rename"]


def resolve_output_path(
    input_path: Path,
    *,
    out_dir: Path | None,
    out_name: str | None,
    policy: PathPolicy | None,
    default_suffix: str,
    default_name_builder: Literal["same_stem", "patched"] = "same_stem",
) -> Path:
    """Build and validate an output path from input and optional overrides."""
    target_dir = out_dir or input_path.parent
    target_dir = policy.ensure_allowed(target_dir) if policy else target_dir.resolve()
    name = normalize_output_name(
        input_path,
        out_name,
        default_suffix=default_suffix,
        default_name_builder=default_name_builder,
    )
    output_path = (target_dir / name).resolve()
    if policy is not None:
        output_path = policy.ensure_allowed(output_path)
    return output_path


def normalize_output_name(
    input_path: Path,
    out_name: str | None,
    *,
    default_suffix: str,
    default_name_builder: Literal["same_stem", "patched"],
) -> str:
    """Normalize output filename with extension fallback behavior."""
    if out_name:
        candidate = Path(out_name)
        return (
            candidate.name if candidate.suffix else f"{candidate.name}{default_suffix}"
        )
    if default_name_builder == "patched":
        return _build_patched_default_name(input_path.stem, default_suffix)
    return f"{input_path.stem}{default_suffix}"


def _build_patched_default_name(stem: str, default_suffix: str) -> str:
    """Build default patched output name without chaining `_patched` repeatedly."""
    if stem.casefold().endswith("_patched"):
        return f"{stem}{default_suffix}"
    return f"{stem}_patched{default_suffix}"


def apply_conflict_policy(
    output_path: Path, on_conflict: OnConflictPolicy
) -> tuple[Path, str | None, bool]:
    """Apply output conflict policy to a resolved output path."""
    if not output_path.exists():
        return output_path, None, False
    if on_conflict == "skip":
        return (
            output_path,
            f"Output exists; skipping write: {output_path.name}",
            True,
        )
    if on_conflict == "rename":
        renamed = next_available_path(output_path)
        return (
            renamed,
            f"Output exists; renamed to: {renamed.name}",
            False,
        )
    return output_path, None, False


def next_available_path(path: Path) -> Path:
    """Return the next available path by appending a numeric suffix."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for idx in range(1, 10_000):
        candidate = path.with_name(f"{stem}_{idx}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Failed to resolve unique path for {path}")
