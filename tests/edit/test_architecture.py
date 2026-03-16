from __future__ import annotations

import ast
from pathlib import Path
import subprocess
import sys

EDIT_DIR = Path(__file__).resolve().parents[2] / "src" / "exstruct" / "edit"


def test_edit_package_has_no_mcp_imports() -> None:
    offenders: list[str] = []
    for path in EDIT_DIR.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    if alias.name.startswith("exstruct.mcp"):
                        offenders.append(f"{path}:{node.lineno}:{alias.name}")
            elif isinstance(node, ast.ImportFrom):
                if node.module and node.module.startswith("exstruct.mcp"):
                    offenders.append(f"{path}:{node.lineno}:{node.module}")
    assert offenders == []


def test_import_exstruct_edit_does_not_load_mcp_package() -> None:
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            "import exstruct.edit, sys; print('exstruct.mcp' in sys.modules)",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    assert result.stdout.strip() == "False"
