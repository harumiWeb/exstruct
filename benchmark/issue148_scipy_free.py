"""Untimed runtime smoke in a genuinely SciPy-free base installation.

Install the base wheel in a clean environment and invoke this module with
--installed to verify packaging as well as runtime behavior. Without that flag,
the module selects the source checkout for development comparisons.
"""

import importlib.util
import argparse
import json
import os
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from unittest.mock import patch


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--installed", action="store_true")
    args = parser.parse_args()
    if importlib.util.find_spec("scipy") is not None:
        raise RuntimeError("This smoke requires an environment without SciPy")
    if importlib.util.find_spec("pandas") is not None:
        raise RuntimeError("This smoke requires an environment without pandas")
    if not args.installed:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
    import numpy as np
    from openpyxl import Workbook
    from openpyxl.styles import Border, Side

    import exstruct
    if args.installed and Path(exstruct.__file__).resolve().is_relative_to(
        Path(__file__).resolve().parents[1] / "src"
    ):
        raise RuntimeError("Installed smoke unexpectedly imported source checkout")
    from exstruct.core import cells

    expected = None
    with TemporaryDirectory(prefix="issue148-base-") as directory:
        path = Path(directory) / "bordered.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        border = Border(bottom=Side(style="thin"))
        for row in range(1, 5):
            for col in range(1, 4):
                sheet.cell(row, col, f"value-{row}-{col}").border = border
        workbook.save(path)
        workbook.close()
        for backend in ("python", "auto", "numpy"):
            with patch.dict(os.environ, {"EXSTRUCT_BORDER_CLUSTER_BACKEND": backend}):
                assert cells.detect_border_clusters(np.ones((2, 2), dtype=bool)) == [
                    (0, 0, 1, 1)
                ]
                result = exstruct.extract(path, mode="light")
                assert result.sheets["Sheet"].rows
                assert result.sheets["Sheet"].table_candidates
                serialized = exstruct.serialize_workbook(result, fmt="json")
                if expected is None:
                    expected = serialized
                assert serialized == expected
    assert not any(name == "scipy" or name.startswith("scipy.") for name in sys.modules)
    print(
        json.dumps(
            {
                "scipy_installed": False,
                "pandas_installed": False,
                "installed_package": args.installed,
                "backends": ["python", "auto", "numpy"],
                "light_output_equal": True,
            }
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
