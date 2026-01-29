from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def sha256_text(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def write_text(p: Path, text: str) -> None:
    ensure_dir(p.parent)
    p.write_text(text, encoding="utf-8")


def write_json(p: Path, obj: Any) -> None:
    ensure_dir(p.parent)
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
