## Plan

- [ ] `tasks/feature_spec.md` のIssue 72仕様を最終確認
- [ ] `src/exstruct/mcp/render_runner.py` を新規追加
- [ ] `src/exstruct/mcp/tools.py` に `CaptureSheetImagesToolInput/Output` と `run_capture_sheet_images_tool` を追加
- [ ] `src/exstruct/mcp/server.py` に `exstruct_capture_sheet_images` ツール登録を追加
- [ ] `src/exstruct/mcp/shared/output_path.py` に画像出力ディレクトリ解決ヘルパを追加
- [ ] `src/exstruct/mcp/shared/a1.py` にシート修飾A1範囲パーサ/正規化を追加
- [ ] `src/exstruct/render/__init__.py` に `sheet` / `a1_range` 指定対応を追加
- [ ] `src/exstruct/mcp/__init__.py` の公開エクスポートを更新
- [ ] `docs/mcp.md` を更新（新ツール仕様、例、制約）
- [ ] `README.md` と `README.ja.md` のMCPツール一覧を更新

## Test Cases

- [ ] `tests/mcp/test_tool_models.py` に新規入力/出力モデル検証を追加
- [ ] `tests/mcp/test_tools_handlers.py` にハンドラ request/result 変換テストを追加
- [ ] `tests/mcp/test_server.py` にツール登録と引数伝播テストを追加
- [ ] `tests/mcp/shared/test_output_path.py` に未指定 `out_dir` 一意化テストを追加
- [ ] `tests/render/test_render_init.py` に `sheet` / `range` 指定出力テストを追加
- [ ] 既存render/mcpテストの回帰確認

## Verification

- [ ] `uv run pytest tests/mcp/test_tool_models.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/shared/test_output_path.py tests/render/test_render_init.py`
- [ ] `uv run task precommit-run`

## Review

- Summary:
- Verification:
- Risks:
- Follow-ups:
