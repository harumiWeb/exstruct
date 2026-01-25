from __future__ import annotations

import json
from typing import Any


def render_markdown(value: Any, *, title: str | None = None) -> str:
    """Render a canonical Markdown representation for JSON-like data.

    Args:
        value: JSON-like payload to render.
        title: Optional top-level title.

    Returns:
        Markdown string representation.
    """
    lines: list[str] = []
    if title:
        lines.append(f"# {title}")
        lines.append("")
    _render_value(lines, value, level=2)
    return "\n".join(lines).strip() + "\n"


def _render_value(lines: list[str], value: Any, *, level: int) -> None:
    """Render a value into Markdown lines.

    Args:
        lines: List to append output lines to.
        value: JSON-like value to render.
        level: Heading level to use for dict sections.
    """
    if isinstance(value, dict):
        _render_dict(lines, value, level=level)
        return
    if isinstance(value, list):
        _render_list(lines, value, level=level)
        return
    lines.append(str(value))


def _render_dict(lines: list[str], value: dict[str, Any], *, level: int) -> None:
    """Render a dict as Markdown sections.

    Args:
        lines: List to append output lines to.
        value: Dict to render.
        level: Heading level for keys.
    """
    for key, item in value.items():
        heading = "#" * max(level, 1)
        lines.append(f"{heading} {key}")
        if isinstance(item, (dict, list)):
            _render_value(lines, item, level=level + 1)
        else:
            lines.append(str(item))
        lines.append("")


def _render_list(lines: list[str], value: list[Any], *, level: int) -> None:
    """Render a list in Markdown.

    Args:
        lines: List to append output lines to.
        value: List to render.
        level: Heading level for nested dicts if needed.
    """
    if not value:
        lines.append("- (empty)")
        return
    if all(isinstance(item, dict) for item in value):
        _render_table(lines, value)
        lines.append("")
        return
    for item in value:
        if isinstance(item, (dict, list)):
            text = _json_string(item)
        else:
            text = str(item)
        lines.append(f"- {text}")


def _render_table(lines: list[str], rows: list[Any]) -> None:
    """Render a list of dicts as a Markdown table.

    Args:
        lines: List to append output lines to.
        rows: List of row dicts.
    """
    keys: list[str] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        for key in row.keys():
            if key not in keys:
                keys.append(key)
    if not keys:
        lines.append("- (empty)")
        return
    header = "| " + " | ".join(keys) + " |"
    sep = "| " + " | ".join(["---"] * len(keys)) + " |"
    lines.append(header)
    lines.append(sep)
    for row in rows:
        if not isinstance(row, dict):
            cells = [_escape_cell(_json_string(row))] + [""] * (len(keys) - 1)
        else:
            cells = [_escape_cell(_cell_value(row.get(k))) for k in keys]
        lines.append("| " + " | ".join(cells) + " |")


def _cell_value(value: Any) -> str:
    """Convert a table cell value to string."""
    if isinstance(value, (dict, list)):
        return _json_string(value)
    if value is None:
        return ""
    return str(value)


def _json_string(value: Any) -> str:
    """Serialize a value as compact JSON for inline use."""
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def _escape_cell(text: str) -> str:
    """Escape pipe characters for Markdown tables."""
    return text.replace("|", "\\|")
