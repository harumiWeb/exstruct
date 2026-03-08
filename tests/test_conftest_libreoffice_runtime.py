"""Tests for LibreOffice runtime gating in ``tests/conftest.py``."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import SimpleNamespace

from _pytest.monkeypatch import MonkeyPatch


def test_has_libreoffice_runtime_returns_false_on_probe_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the smoke gate rejects incompatible bridge runtimes."""

    module_path = Path(__file__).with_name("conftest.py")
    spec = importlib.util.spec_from_file_location("tests_runtime_conftest", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Failed to load tests/conftest.py")
    conftest = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(conftest)

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")

    def _raise_probe_failure(_path: Path) -> Path:
        """Simulate an explicit probe compatibility failure."""

        raise RuntimeError("bridge probe failed")

    monkeypatch.setattr(
        conftest,
        "_load_libreoffice_runtime_module",
        lambda: SimpleNamespace(
            _which_soffice=lambda: soffice_path,
            _resolve_python_path=_raise_probe_failure,
        ),
    )
    monkeypatch.delenv("EXSTRUCT_LIBREOFFICE_PATH", raising=False)
    conftest._has_libreoffice_runtime.cache_clear()

    assert conftest._has_libreoffice_runtime() is False

    conftest._has_libreoffice_runtime.cache_clear()
