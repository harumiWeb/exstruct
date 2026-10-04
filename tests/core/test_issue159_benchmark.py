"""No-Excel regressions for the Issue 159 benchmark metrics."""

import logging

from _pytest.monkeypatch import MonkeyPatch
from benchmark import issue159_colors


def test_measure_preserves_legacy_counter_and_omits_hybrid_false_zero(
    monkeypatch: MonkeyPatch,
) -> None:
    """Keep the legacy counter while omitting the hybrid per-sheet false zero."""
    counter = issue159_colors.cells._display_format_calls
    active_strategy = "reference"

    def fake_display_format_color(_sheet: object, _row: int, _col: int) -> int:
        if active_strategy == "reference":
            counter.set(counter.get() + 1)
        return 255

    def fake_reference_extract(
        _book: object,
        _include_default_background: bool,
        _ignore_colors: set[str] | None,
    ) -> None:
        issue159_colors.cells._get_display_format_color(None, 1, 1)

    def fake_hybrid_extract(
        _book: object,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> None:
        issue159_colors.cells._get_display_format_color(None, 1, 1)
        logging.getLogger("exstruct.core.color_hybrid").debug(
            "synthetic hybrid measurement",
            extra={"display_format_calls": 1},
        )
        # The real hybrid extractor resets this per-sheet counter after logging.
        counter.set(0)

    monkeypatch.setattr(
        issue159_colors.cells, "_get_display_format_color", fake_display_format_color
    )
    monkeypatch.setattr(issue159_colors, "_reference_extract", fake_reference_extract)
    monkeypatch.setattr(
        issue159_colors.cells, "extract_sheet_colors_map_com", fake_hybrid_extract
    )

    outer_token = counter.set(37)
    try:
        legacy_record, _ = issue159_colors._measure(
            object(), "reference", False, None, "static", 1
        )

        assert counter.get() == 37
        active_strategy = "hybrid"
        hybrid_record, _ = issue159_colors._measure(
            object(), "hybrid", False, None, "static", 1
        )

        assert counter.get() == 37
        assert legacy_record["wrapper_display_format_calls"] == 1
        assert legacy_record["cells_context_display_format_calls"] == 1
        assert hybrid_record["wrapper_display_format_calls"] == 1
        assert "cells_context_display_format_calls" not in hybrid_record
        assert (
            hybrid_record["hybrid_debug_records"][0]["metric_fields"][
                "display_format_calls"
            ]
            == 1
        )
    finally:
        counter.reset(outer_token)
