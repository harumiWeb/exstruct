## Plan

- [x] `tasks/feature_spec.md` のIssue 72仕様を最終確認
- [x] `src/exstruct/mcp/render_runner.py` を新規追加
- [x] `src/exstruct/mcp/tools.py` に `CaptureSheetImagesToolInput/Output` と `run_capture_sheet_images_tool` を追加
- [x] `src/exstruct/mcp/server.py` に `exstruct_capture_sheet_images` ツール登録を追加
- [x] `src/exstruct/mcp/shared/output_path.py` に画像出力ディレクトリ解決ヘルパを追加
- [x] `src/exstruct/mcp/shared/a1.py` にシート修飾A1範囲パーサ/正規化を追加
- [x] `src/exstruct/render/__init__.py` に `sheet` / `a1_range` 指定対応を追加
- [x] `src/exstruct/mcp/__init__.py` の公開エクスポートを更新
- [x] `docs/mcp.md` を更新（新ツール仕様、例、制約）
- [x] `README.md` と `README.ja.md` のMCPツール一覧を更新

## Test Cases

- [x] `tests/mcp/test_tool_models.py` に新規入力/出力モデル検証を追加
- [x] `tests/mcp/test_tools_handlers.py` にハンドラ request/result 変換テストを追加
- [x] `tests/mcp/test_server.py` にツール登録と引数伝播テストを追加
- [x] `tests/mcp/shared/test_output_path.py` に未指定 `out_dir` 一意化テストを追加
- [x] `tests/render/test_render_init.py` に `sheet` / `range` 指定出力テストを追加
- [x] 既存render/mcpテストの回帰確認

## Verification

- [x] `uv run pytest tests/mcp/test_tool_models.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/shared/test_output_path.py tests/render/test_render_init.py`
- [x] `uv run task precommit-run`

## Review

- Summary:
  - Issue 72として `exstruct_capture_sheet_images` をMCPに追加し、`sheet` / `range` 指定、`out_dir` 未指定時の一意ディレクトリ解決、COM必須チェックを実装した。
  - `render.export_sheet_images` に `sheet` / `a1_range` を追加し、対象シート/範囲のみ出力できるようにした。
- Verification:
  - `uv run pytest tests/mcp/test_tool_models.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/shared/test_output_path.py tests/render/test_render_init.py` -> 148 passed
  - `uv run task precommit-run` -> ruff / ruff-format / mypy passed
- Risks:
  - `sheet` 名一致は厳密一致（大文字小文字含む）で比較しているため、利用側が異なる表記を渡すと不一致エラーになる。
  - 範囲指定時のレンダリング結果ページ数はExcelの改ページ設定に依存する。
- Follow-ups:
  - 必要であれば `sheet` 名の大小無視比較オプションを検討する。
  - 将来的に `copyPicture` 以外の軽量レンダリング経路を比較評価する。
