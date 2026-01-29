from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel


class RenderConfig(BaseModel):
    dpi: int = 200
    max_pages: int = 6


class Case(BaseModel):
    id: str
    type: str
    xlsx: str
    question: str
    truth: str
    sheet_scope: list[str] | None = None
    render: RenderConfig = RenderConfig()


class Manifest(BaseModel):
    cases: list[Case]


def load_manifest(path: Path) -> Manifest:
    data = json.loads(path.read_text(encoding="utf-8"))
    return Manifest(**data)
