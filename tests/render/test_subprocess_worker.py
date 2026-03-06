from __future__ import annotations

import json
from pathlib import Path

import pytest

from exstruct.render import subprocess_worker


def _build_request_payload(
    tmp_path: Path, *, result_path: Path | None = None
) -> dict[str, object]:
    """Build a minimally valid worker request payload for tests."""
    return {
        "pdf_path": str(tmp_path / "sheet.pdf"),
        "output_dir": str(tmp_path / "images"),
        "sheet_index": 0,
        "safe_name": "Sheet1",
        "dpi": 144,
        "started_path": str(tmp_path / "started.txt"),
        "result_path": str(result_path or (tmp_path / "result.json")),
    }


def test_main_writes_success_payload(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Write success payload and startup marker when rendering succeeds."""
    result_path = tmp_path / "result.json"
    payload = _build_request_payload(tmp_path, result_path=result_path)
    request_file = tmp_path / "request.json"
    request_file.write_text(json.dumps(payload), encoding="utf-8")

    monkeypatch.setattr(
        subprocess_worker,
        "_render_pdf_pages",
        lambda request: [str(Path(request.output_dir) / "01_Sheet1.png")],
    )

    exit_code = subprocess_worker.main(["--request-file", str(request_file)])

    assert exit_code == 0
    assert Path(str(payload["started_path"])).exists()
    assert json.loads(result_path.read_text(encoding="utf-8")) == {
        "paths": [str(tmp_path / "images" / "01_Sheet1.png")]
    }


def test_main_writes_failure_payload_on_render_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Write failure payload when rendering fails after request load."""
    result_path = tmp_path / "result.json"
    payload = _build_request_payload(tmp_path, result_path=result_path)
    request_file = tmp_path / "request.json"
    request_file.write_text(json.dumps(payload), encoding="utf-8")

    def _raise_render(_: subprocess_worker.RenderWorkerRequest) -> list[str]:
        raise RuntimeError("boom")

    monkeypatch.setattr(subprocess_worker, "_render_pdf_pages", _raise_render)

    exit_code = subprocess_worker.main(["--request-file", str(request_file)])

    assert exit_code == 1
    error_payload = json.loads(result_path.read_text(encoding="utf-8"))
    assert error_payload["error"] == "RuntimeError: boom"


def test_main_preserves_actionable_error_before_request_load(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    """Emit diagnostics and best-effort error payload before full request parsing."""
    result_path = tmp_path / "result.json"
    payload = _build_request_payload(tmp_path, result_path=result_path)
    payload["sheet_index"] = "not-an-int"
    request_file = tmp_path / "request.json"
    request_file.write_text(json.dumps(payload), encoding="utf-8")

    exit_code = subprocess_worker.main(["--request-file", str(request_file)])

    captured = capsys.readouterr()
    assert exit_code == 1
    assert "ValueError" in captured.err
    error_payload = json.loads(result_path.read_text(encoding="utf-8"))
    assert "ValueError" in error_payload["error"]
