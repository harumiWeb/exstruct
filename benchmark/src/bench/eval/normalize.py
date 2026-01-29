from __future__ import annotations

import json
import re
from typing import Any


def _strip_code_fences(s: str) -> str:
    s = s.strip()
    s = re.sub(r"^```(json)?\s*", "", s)
    s = re.sub(r"\s*```$", "", s)
    return s.strip()


def normalize_json_text(s: str) -> Any:
    """
    LLM出力を JSON として読み、正規化されたPythonオブジェクトを返す
    """
    s = _strip_code_fences(s)
    # 余計な前後テキストが入った場合の救済（最初の{...}を拾う）
    if "{" in s and "}" in s:
        start = s.find("{")
        end = s.rfind("}")
        s = s[start : end + 1]
    obj = json.loads(s)
    return obj
