"""macOS keeps LibreOffice's bundled Python outside the soffice executable's dir."""

from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.core import libreoffice


def _fake_bundle(root: Path) -> tuple[Path, Path]:
    """Build a minimal LibreOffice.app layout and return (soffice, python)."""
    macos_dir = root / "LibreOffice.app" / "Contents" / "MacOS"
    resources_dir = root / "LibreOffice.app" / "Contents" / "Resources"
    macos_dir.mkdir(parents=True)
    resources_dir.mkdir(parents=True)
    soffice = macos_dir / "soffice"
    python = resources_dir / "python"
    soffice.write_text("#!/bin/sh\n")
    python.write_text("#!/bin/sh\n")
    return soffice, python


def test_program_dirs_include_the_sibling_resources_dir(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(libreoffice.sys, "platform", "darwin")
    soffice, python = _fake_bundle(tmp_path)

    dirs = libreoffice._soffice_program_dirs(soffice)

    assert python.parent in dirs
    candidates = [c for d in dirs for c in libreoffice._bundled_python_candidates(d)]
    assert python in candidates


def test_program_dirs_fall_back_to_the_bundle_for_a_homebrew_wrapper(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Homebrew puts a bash wrapper on PATH, so the resolved path leaves the bundle."""
    monkeypatch.setattr(libreoffice.sys, "platform", "darwin")
    _, python = _fake_bundle(tmp_path)
    monkeypatch.setattr(
        libreoffice, "_MACOS_BUNDLE_MACOS_DIR", python.parent.parent / "MacOS"
    )
    wrapper_dir = tmp_path / "command-wrappers"
    wrapper_dir.mkdir()
    wrapper = wrapper_dir / "soffice"
    wrapper.write_text("#!/bin/bash\n")

    dirs = libreoffice._soffice_program_dirs(wrapper)

    assert python.parent in dirs


def test_program_dirs_stay_unchanged_off_macos(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(libreoffice.sys, "platform", "linux")
    soffice, python = _fake_bundle(tmp_path)

    assert python.parent not in libreoffice._soffice_program_dirs(soffice)
