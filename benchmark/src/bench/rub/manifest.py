from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel, Field


class RubTask(BaseModel):
    """RUB task definition."""

    id: str
    track: str = Field("reconstruction", description="Evaluation track name.")
    source_case_id: str = Field(..., description="Case id for Stage A Markdown.")
    type: str
    question: str
    truth: str
    schema_path: str | None = None
    unordered_paths: list[str] | None = None


class RubManifest(BaseModel):
    """RUB manifest container."""

    tasks: list[RubTask]


def load_rub_manifest(path: Path) -> RubManifest:
    """Load a RUB manifest file.

    Args:
        path: Path to rub/manifest.json.

    Returns:
        Parsed RubManifest.
    """
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    return RubManifest(**data)
