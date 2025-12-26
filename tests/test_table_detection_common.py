from openpyxl.utils import get_column_letter

from exstruct.core.cells import (
    _collect_table_candidates_from_values,
    _merge_rectangles,
)


def test_merge_rectangles_keeps_contained_rects() -> None:
    rects = [(1, 1, 3, 3), (2, 2, 2, 2)]
    merged = _merge_rectangles(rects)
    assert len(merged) == 2


def test_collect_table_candidates_builds_range() -> None:
    values = [[1, 2], [3, 4]]
    candidates = _collect_table_candidates_from_values(
        values,
        base_top=1,
        base_left=1,
        col_name=get_column_letter,
    )
    assert candidates == ["A1:B2"]
