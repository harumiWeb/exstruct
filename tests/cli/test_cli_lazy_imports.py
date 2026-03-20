from __future__ import annotations

import json
import subprocess
import sys
from typing import cast


def _run_python(code: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-c", code],
        check=False,
        capture_output=True,
        text=True,
    )


def _run_probe(code: str) -> dict[str, object]:
    result = _run_python(code)
    assert result.returncode == 0, result.stderr or result.stdout
    return cast(dict[str, object], json.loads(result.stdout))


def test_import_exstruct_stays_lightweight() -> None:
    payload = _run_probe(
        """
import json
import sys
import exstruct
print(json.dumps({
    "core_integrate": "exstruct.core.integrate" in sys.modules,
    "engine": "exstruct.engine" in sys.modules,
    "pandas": "pandas" in sys.modules,
    "openpyxl": "openpyxl" in sys.modules,
    "xlwings": "xlwings" in sys.modules,
}))
"""
    )

    assert payload == {
        "core_integrate": False,
        "engine": False,
        "pandas": False,
        "openpyxl": False,
        "xlwings": False,
    }


def test_import_cli_main_does_not_load_edit_or_extraction_modules() -> None:
    payload = _run_probe(
        """
import json
import sys
import exstruct.cli.main
print(json.dumps({
    "cli_edit": "exstruct.cli.edit" in sys.modules,
    "mcp": "exstruct.mcp" in sys.modules,
    "core_integrate": "exstruct.core.integrate" in sys.modules,
    "engine": "exstruct.engine" in sys.modules,
}))
"""
    )

    assert payload == {
        "cli_edit": False,
        "mcp": False,
        "core_integrate": False,
        "engine": False,
    }


def test_import_cli_edit_does_not_load_mcp_or_edit_execution_modules() -> None:
    payload = _run_probe(
        """
import json
import sys
import exstruct.cli.edit
print(json.dumps({
    "mcp": "exstruct.mcp" in sys.modules,
    "extract_runner": "exstruct.mcp.extract_runner" in sys.modules,
    "edit_api": "exstruct.edit.api" in sys.modules,
    "edit_service": "exstruct.edit.service" in sys.modules,
}))
"""
    )

    assert payload == {
        "mcp": False,
        "extract_runner": False,
        "edit_api": False,
        "edit_service": False,
    }


def test_help_and_ops_list_keep_lightweight_import_boundaries() -> None:
    payload = _run_probe(
        """
import io
import json
import sys
from contextlib import redirect_stderr, redirect_stdout

from exstruct.cli.main import main

stdout_buffer = io.StringIO()
stderr_buffer = io.StringIO()
with redirect_stdout(stdout_buffer), redirect_stderr(stderr_buffer):
    try:
        help_exit = main(["--help"])
    except SystemExit as exc:
        help_exit = exc.code
help_text = stdout_buffer.getvalue()

stdout_buffer = io.StringIO()
stderr_buffer = io.StringIO()
with redirect_stdout(stdout_buffer), redirect_stderr(stderr_buffer):
    ops_exit = main(["ops", "list"])
ops_payload = json.loads(stdout_buffer.getvalue())

print(json.dumps({
    "help_exit": help_exit,
    "help_has_option": "--auto-page-breaks-dir" in help_text,
    "ops_exit": ops_exit,
    "ops_has_set_value": any(item["op"] == "set_value" for item in ops_payload["ops"]),
    "mcp": "exstruct.mcp" in sys.modules,
    "extract_runner": "exstruct.mcp.extract_runner" in sys.modules,
    "core_integrate": "exstruct.core.integrate" in sys.modules,
}))
"""
    )

    assert payload["help_exit"] == 0
    assert payload["help_has_option"] is True
    assert payload["ops_exit"] == 0
    assert payload["ops_has_set_value"] is True
    assert payload["mcp"] is False
    assert payload["extract_runner"] is False
    assert payload["core_integrate"] is False
