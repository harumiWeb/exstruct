"""Clustering parity and SciPy-free runtime contracts for issue 148."""

from importlib.metadata import PackageNotFoundError
import json
from pathlib import Path
import subprocess
import sys
from unittest.mock import Mock

from benchmark import issue148_clustering as comparison
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Border, Side
import pytest

import exstruct
from exstruct.core import cells


@pytest.mark.parametrize("min_size", [-1, 0, 1, 2, 4, 8, 100])
def test_cluster_parity(min_size: int) -> None:
    """Compare exact order and bboxes, including holes and random maps."""
    grids = [
        np.zeros((0, 4), dtype=bool),
        np.zeros((4, 0), dtype=bool),
        np.zeros((3, 7), dtype=bool),
        np.ones((5, 8), dtype=bool),
        np.eye(8, dtype=bool),
        np.array([[1, 0, 1, 1], [1, 0, 0, 0], [0, 1, 1, 1]], dtype=bool),
        np.array([[1, 1, 1], [1, 0, 1], [1, 1, 1]], dtype=bool),
    ]
    random = np.random.default_rng(148)
    grids.extend(random.random((13, 19)) < density for density in (0.05, 0.3, 0.8))
    grids.extend([grids[-1].T, grids[-1][::-1, ::-1]])
    for grid in grids:
        original = grid.copy()
        expected = cells._detect_border_clusters_numpy(grid, min_size)
        assert cells._detect_border_clusters_python(grid, min_size) == expected
        assert cells._detect_border_clusters_sparse(grid, min_size) == expected
        np.testing.assert_array_equal(grid, original)


def test_component_size_and_row_major_seed_order() -> None:
    """Four neighbors, occupied-cell threshold and seed order are independent."""
    grid = np.array([[0, 0, 1, 1], [1, 0, 0, 0], [1, 1, 1, 0]], dtype=bool)
    # Second seed's bbox starts left of the first: do not sort boxes lexically.
    for implementation in (
        cells._detect_border_clusters_numpy,
        cells._detect_border_clusters_python,
        cells._detect_border_clusters_sparse,
    ):
        assert implementation(grid, 2) == [(0, 2, 0, 3), (1, 0, 2, 2)]
        assert implementation(grid, 4) == [(1, 0, 2, 2)]
        assert implementation(np.eye(4, dtype=bool), 2) == []


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("", "auto"),
        ("auto", "auto"),
        (" PYTHON ", "python"),
        ("NumPy", "numpy"),
        ("scipy", "auto"),
        ("invalid", "auto"),
    ],
)
def test_backend_names_unchanged(
    monkeypatch: pytest.MonkeyPatch, value: str, expected: str
) -> None:
    monkeypatch.setenv("EXSTRUCT_BORDER_CLUSTER_BACKEND", value)
    assert cells._resolve_border_cluster_backend() == expected


@pytest.mark.parametrize("backend", ["auto", "numpy"])
@pytest.mark.parametrize("error", [ImportError("blocked scipy"), RuntimeError("label")])
def test_accelerated_failure_and_live_overrides(
    monkeypatch: pytest.MonkeyPatch, backend: str, error: Exception
) -> None:
    grid = np.ones((2, 2), dtype=bool)
    accelerated = Mock(side_effect=error)
    fallback = Mock(return_value=[(9, 8, 7, 6)])
    monkeypatch.setenv("EXSTRUCT_BORDER_CLUSTER_BACKEND", backend)
    monkeypatch.setattr(cells, "_detect_border_clusters_numpy", accelerated)
    monkeypatch.setattr(cells, "_detect_border_clusters_python", fallback)
    assert cells.detect_border_clusters(grid, 3) == [(9, 8, 7, 6)]
    accelerated.assert_called_once_with(grid, 3)
    fallback.assert_called_once_with(grid, 3)


def test_explicit_python_skips_accelerated_backend(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("EXSTRUCT_BORDER_CLUSTER_BACKEND", "python")
    accelerated = Mock(side_effect=AssertionError("must not call scipy"))
    monkeypatch.setattr(cells, "_detect_border_clusters_numpy", accelerated)
    assert cells.detect_border_clusters(np.ones((2, 2), dtype=bool)) == [(0, 0, 1, 1)]
    accelerated.assert_not_called()


@pytest.mark.parametrize("backend", ["python", "auto", "numpy"])
def test_light_extraction_with_scipy_imports_blocked(
    tmp_path: Path, backend: str
) -> None:
    """Block SciPy before importing ExStruct in a fresh child interpreter."""
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    border = Border(bottom=Side(style="thin"))
    for row in range(1, 5):
        for col in range(1, 4):
            sheet.cell(row, col, f"value-{row}-{col}").border = border
    path = tmp_path / "bordered.xlsx"
    workbook.save(path)
    workbook.close()
    script = """
import importlib.abc
import os
import sys
class NoScipy(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname == 'scipy' or fullname.startswith('scipy.'):
            raise ModuleNotFoundError('SciPy blocked for issue 148')
sys.meta_path.insert(0, NoScipy())
os.environ['EXSTRUCT_BORDER_CLUSTER_BACKEND'] = sys.argv[2]
import exstruct
from exstruct.core import cells
import numpy as np
assert cells.detect_border_clusters(np.ones((2, 2), dtype=bool)) == [(0, 0, 1, 1)]
result = exstruct.extract(sys.argv[1], mode='light')
assert result.sheets['Sheet'].table_candidates
assert result.sheets['Sheet'].rows
assert not any(name == 'scipy' or name.startswith('scipy.') for name in sys.modules)
"""
    child = subprocess.run(
        [sys.executable, "-c", script, str(path), backend],
        capture_output=True,
        text=True,
        timeout=90,
        check=False,
    )
    assert child.returncode == 0, child.stdout + child.stderr


def test_comparison_optional_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    """A missing optional pandas installation must not block comparison."""
    monkeypatch.setattr(comparison, "version", Mock(side_effect=PackageNotFoundError))
    assert comparison.dependency_version("pandas") is None
    monkeypatch.setattr(comparison, "version", Mock(return_value="1.2.3"))
    assert comparison.dependency_version("numpy") == "1.2.3"


@pytest.mark.parametrize(
    "error", [FileNotFoundError("git"), subprocess.TimeoutExpired("git", 10)]
)
def test_comparison_git_metadata_unavailable(
    monkeypatch: pytest.MonkeyPatch, error: Exception
) -> None:
    monkeypatch.setattr(subprocess, "run", Mock(side_effect=error))
    assert comparison.git_metadata(["rev-parse", "HEAD"]) is None


def test_comparison_check_only_never_times(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(comparison, "perf_counter", Mock(side_effect=AssertionError))
    assert comparison.observations(lambda: "result", 7, True) == ("result", {})


def test_comparison_main_preserves_unknown_metadata(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    output = tmp_path / "comparison.json"
    monkeypatch.setattr(
        sys, "argv", ["comparison", "--check-only", "--output", str(output)]
    )
    monkeypatch.setattr(comparison, "git_metadata", Mock(return_value=None))
    monkeypatch.setattr(comparison, "dependency_version", Mock(return_value=None))
    monkeypatch.setattr(comparison, "run_comparison", Mock(return_value={}))
    assert comparison.main() == 0
    result = json.loads(output.read_text(encoding="utf-8"))
    assert result["source_revision"] is None
    assert result["source_dirty"] is None
    assert result["dependencies"]["pandas"] is None
    assert result["warm_repeats"] == 0


def test_comparison_extraction_restores_live_override(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The evaluator restores monkeypatch/environment surfaces on failure."""
    original = cells._detect_border_clusters_python
    monkeypatch.setenv("EXSTRUCT_BORDER_CLUSTER_BACKEND", "auto")
    monkeypatch.setattr(exstruct, "extract", Mock(side_effect=ValueError))
    with pytest.raises(ValueError):
        comparison.extraction(
            Path("not-used.xlsx"), cells._detect_border_clusters_sparse
        )
    assert cells._detect_border_clusters_python is original
    assert cells._resolve_border_cluster_backend() == "auto"
