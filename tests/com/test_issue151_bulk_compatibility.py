"""Real Excel evidence for the experimental cell reader's adoption gate."""

from benchmark.issue151_compatibility import probe
import pytest


@pytest.mark.com
def test_bulk_com_compatibility_gate() -> None:
    result = probe()
    assert result["current_com_succeeded"]
    assert result["candidate_com_succeeded"]
    assert not result["full_output_equal"]
    saved = result["file_cells"]["Sheet"]
    live = result["com_cells"]["Sheet"]
    unsaved = result["unsaved_com_cells"]["Sheet"]
    assert saved[1]["c"] == {"1": "2024-01-02 03:04:05"}
    assert isinstance(live[1]["c"]["1"], float)
    assert not any(row["r"] == 3 for row in saved)
    assert next(row for row in live if row["r"] == 3)["c"] == {"2": 3}
    assert unsaved[0]["c"] == {"0": "unsaved"}
    for rows in (saved, live):
        assert not any(row["r"] == 4 for row in rows)
        assert next(row for row in rows if row["r"] == 5)["c"] == {"4": -2146826281}
        assert next(row for row in rows if row["r"] == 7)["links"] == {
            "6": "https://example.test/"
        }
        assert next(row for row in rows if row["r"] == 8)["links"] is None
