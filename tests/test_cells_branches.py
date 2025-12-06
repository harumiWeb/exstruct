import numpy as np

from exstruct.core import cells


def test_table_density_metrics() -> None:
    matrix = [
        ["a", ""],
        ["", "b"],
    ]
    density, coverage = cells._table_density_metrics(matrix)
    # 2 nonempty / 4 total = 0.5
    assert density == 0.5
    # bounding box covers all cells here => coverage 1.0
    assert coverage == 1.0


def test_is_plausible_table_rejects_too_small() -> None:
    assert cells._is_plausible_table([["only one row"]]) is False


def test_table_signal_score_increases_with_header() -> None:
    matrix = [
        ["Name", "Age"],  # header-like row (strings dominate)
        ["Alice", "30"],
        ["Bob", "25"],
    ]
    with_header = cells._table_signal_score(matrix)
    # no-header case: all numeric-like values → header bonusなし
    no_header = cells._table_signal_score([["1", "2"], ["3", "4"]])
    assert with_header > no_header


def test_nonempty_clusters_two_components() -> None:
    matrix = [
        ["x", ""],
        ["", ""],
        ["", "y"],
    ]
    boxes = cells._nonempty_clusters(matrix)
    assert (0, 0, 0, 0) in boxes
    assert (2, 1, 2, 1) in boxes


def test_detect_border_clusters_fallback() -> None:
    has_border = np.array(
        [
            [False, True, False],
            [True, True, False],
            [False, False, False],
        ],
        dtype=bool,
    )
    rects = cells.detect_border_clusters(has_border, min_size=2)
    # single cluster covering the three True cells -> bbox (0,0)-(1,1)
    assert (0, 0, 1, 1) in rects
