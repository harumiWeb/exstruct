from __future__ import annotations

from pathlib import Path

from pydantic import BaseModel, Field


class PathPolicy(BaseModel):
    """Filesystem access policy for MCP requests."""

    root: Path = Field(..., description="Root directory for allowed access.")
    deny_globs: list[str] = Field(
        default_factory=list, description="Glob patterns to deny."
    )

    def normalize_root(self) -> Path:
        """Return the resolved root path.

        Returns:
            Resolved root directory path.
        """
        return self.root.resolve()

    def ensure_allowed(self, path: Path) -> Path:
        """Validate that a path is within root and not denied.

        Args:
            path: Candidate path to validate.

        Returns:
            Resolved path if allowed.

        Raises:
            ValueError: If the path is outside the root or denied by glob.
        """
        resolved = path.resolve()
        root = self.normalize_root()
        if resolved != root and root not in resolved.parents:
            raise ValueError(f"Path is outside root: {resolved}")
        if self._is_denied(resolved, root):
            raise ValueError(f"Path is denied by policy: {resolved}")
        return resolved

    def _is_denied(self, path: Path, root: Path) -> bool:
        """Check if a path is denied by glob rules.

        Args:
            path: Candidate path to check.
            root: Resolved root path.

        Returns:
            True if denied, False otherwise.
        """
        try:
            rel = path.relative_to(root)
        except ValueError:
            return True
        for pattern in self.deny_globs:
            if rel.match(pattern) or path.match(pattern):
                return True
        return False
