import shutil
import subprocess
import sys
from pathlib import Path

import pytest

import xlwings as xw


def _excel_available() -> bool:
    try:
        app = xw.App(add_book=False, visible=False)
        app.quit()
        return True
    except Exception:
        return False


def _copy_sample_excel(tmp_path: Path) -> Path:
    sample = Path("sample") / "sample.xlsx"
    dest = tmp_path / "sample.xlsx"
    shutil.copy(sample, dest)
    return dest


def test_CLIでjson出力が成功する(tmp_path: Path) -> None:
    xlsx = _copy_sample_excel(tmp_path)
    out_json = tmp_path / "out.json"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(xlsx), "-o", str(out_json)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert out_json.exists()
    assert f"{xlsx.name} -> {out_json}" in result.stdout


def test_CLIでyamlやtoon指定は未サポート(tmp_path: Path) -> None:
    xlsx = _copy_sample_excel(tmp_path)
    out_yaml = tmp_path / "out.yaml"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(xlsx), "-o", str(out_yaml), "-f", "yaml"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert out_yaml.exists()
    out_toon = tmp_path / "out.toon"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(xlsx), "-o", str(out_toon), "-f", "toon"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert out_toon.exists()


@pytest.mark.skipif(not _excel_available(), reason="Excel COM unavailable; skipping PDF/PNG export tests.")
def test_CLIでpdfと画像が出力される(tmp_path: Path) -> None:
    xlsx = _copy_sample_excel(tmp_path)
    out_json = tmp_path / "out.json"
    cmd = [
        sys.executable,
        "-m",
        "exstruct.cli.main",
        str(xlsx),
        "-o",
        str(out_json),
        "--pdf",
        "--image",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    pdf_path = out_json.with_suffix(".pdf")
    images_dir = out_json.parent / f"{out_json.stem}_images"
    assert pdf_path.exists()
    assert images_dir.exists()
    assert any(images_dir.glob("*.png"))


def test_CLIで無効ファイルは安全終了する(tmp_path: Path) -> None:
    bad_path = tmp_path / "nope.xlsx"
    out_json = tmp_path / "out.json"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(bad_path), "-o", str(out_json)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert "not found" in (result.stdout + result.stderr).lower()
