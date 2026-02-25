from __future__ import annotations

import pytest

from exstruct.mcp.patch.normalize import (
    coerce_patch_ops,
    resolve_top_level_sheet_for_payload,
)


def test_coerce_patch_ops_normalizes_aliases() -> None:
    result = coerce_patch_ops(
        [
            {"op": "add_sheet", "name": "Data"},
            {
                "op": "set_dimensions",
                "sheet": "Data",
                "col": ["A", 2],
                "width": 18,
                "row": [1],
                "height": 24,
            },
            {
                "op": "set_alignment",
                "sheet": "Data",
                "cell": "A1",
                "horizontal": "center",
                "vertical": "bottom",
            },
            {
                "op": "set_fill_color",
                "sheet": "Data",
                "cell": "B1",
                "color": "#D9E1F2",
            },
        ]
    )
    assert result[0] == {"op": "add_sheet", "sheet": "Data"}
    assert result[1]["columns"] == ["A", 2]
    assert result[1]["column_width"] == 18
    assert result[1]["rows"] == [1]
    assert result[1]["row_height"] == 24
    assert result[2]["horizontal_align"] == "center"
    assert result[2]["vertical_align"] == "bottom"
    assert result[3]["fill_color"] == "#D9E1F2"


def test_coerce_patch_ops_normalizes_draw_grid_border_range() -> None:
    result = coerce_patch_ops(
        [{"op": "draw_grid_border", "sheet": "Sheet1", "range": "B3:D5"}]
    )
    assert result[0]["base_cell"] == "B3"
    assert result[0]["row_count"] == 3
    assert result[0]["col_count"] == 3
    assert "range" not in result[0]


def test_resolve_top_level_sheet_for_payload() -> None:
    payload = {
        "sheet": "Sheet1",
        "ops": [
            {"op": "set_value", "cell": "A1", "value": "x"},
            {"op": "add_sheet", "name": "Data"},
        ],
    }
    resolved = resolve_top_level_sheet_for_payload(payload)
    assert isinstance(resolved, dict)
    ops = resolved["ops"]
    assert ops[0]["sheet"] == "Sheet1"
    assert ops[1]["sheet"] == "Data"


def test_resolve_top_level_sheet_for_payload_rejects_missing_sheet() -> None:
    with pytest.raises(ValueError, match="missing sheet"):
        resolve_top_level_sheet_for_payload(
            {"ops": [{"op": "set_value", "cell": "A1", "value": "x"}]}
        )
