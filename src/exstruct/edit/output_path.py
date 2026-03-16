"""Output-path helpers owned by the public workbook editing core."""

from __future__ import annotations

from pathlib import Path
from typing import Literal, Protocol

from .types import OnConflictPolicy


class PathPolicyProtocol(Protocol):
    """Structural protocol for host-owned path policy objects."""

    def ensure_allowed(self, path: Path) -> Path:
        """Resolve and validate a path against host policy."""
        ...

    def normalize_root(self) -> Path:
        """Return the normalized root path for host policy."""
        ...


def resolve_output_path(
    input_path: Path,
    *,
    out_dir: Path | None,
    out_name: str | None,
    policy: PathPolicyProtocol | None,
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


def resolve_image_output_dir(
    input_path: Path,
    *,
    out_dir: Path | None,
    policy: PathPolicyProtocol | None,
) -> Path:
    """Resolve output directory for sheet image export.

    If `out_dir` is omitted, a unique `<stem>_images` directory is created
    under MCP root (or input parent when policy is not provided).
    """
    if out_dir is not None:
        return policy.ensure_allowed(out_dir) if policy else out_dir.resolve()
    base_dir = policy.normalize_root() if policy else input_path.parent.resolve()
    candidate = (base_dir / f"{input_path.stem}_images").resolve()
    if policy is not None:
        candidate = policy.ensure_allowed(candidate)
    return next_available_directory(candidate, policy=policy)


def next_available_directory(path: Path, *, policy: PathPolicyProtocol | None) -> Path:
    """Reserve and return a unique directory path with numeric suffix when needed."""
    reserved = _reserve_directory(path)
    if reserved is not None:
        if policy is not None:
            policy.ensure_allowed(reserved)
        return reserved
    stem = path.name
    for idx in range(1, 10_000):
        candidate = path.with_name(f"{stem}_{idx}")
        if policy is not None:
            candidate = policy.ensure_allowed(candidate)
        reserved = _reserve_directory(candidate)
        if reserved is not None:
            return reserved
    raise RuntimeError(f"Failed to resolve unique path for {path}")


def _reserve_directory(path: Path) -> Path | None:
    """Create one directory atomically and return path when successful."""
    try:
        path.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        return None
    return path.resolve()
