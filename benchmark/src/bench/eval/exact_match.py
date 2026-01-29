from __future__ import annotations

import json
from typing import Any


def canonical(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def exact_match(a: Any, b: Any) -> bool:
    return canonical(a) == canonical(b)
