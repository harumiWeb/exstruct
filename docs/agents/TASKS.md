# Task List

未完了 [ ], 完了 [x]

- [x] `FEATURE_SPEC.md` に exstruct_patch の詳細仕様（I/F、検証、差分、エラー）を記載
- [x] `mcp/patch_runner.py` を新設し、`PatchRequest/Result` を定義
- [x] `mcp/tools.py` に `PatchToolInput/Output` と `run_patch_tool` を追加
- [x] `mcp/server.py` に `exstruct_patch` 登録を追加
- [x] パス制御・競合ポリシー・保存ロジックを実装
- [x] openpyxlベースの適用処理を実装
- [x] COM利用時の分岐（可能なら）とwarning出力
- [x] ユニットテスト（非COM）を追加

## レビュー対応タスク（MVP修正）

- [x] `mcp/server.py`: `exstruct_patch` の `on_conflict` 解決で固定 `rename` を廃止し、サーバー既定値（`default_on_conflict`）を利用する
- [x] `mcp/patch_runner.py`: 保存前に出力先ディレクトリ作成処理を追加する（未作成 `out_dir` 対応）
- [x] `mcp/patch_runner.py`: `.xls` + COM不可時のエラーメッセージを「COM必須」が明確な文言に調整する
- [x] `docs/agents/FEATURE_SPEC.md`: 拡張子要件を「`.xls` はCOM利用可能環境のみ」に更新する（本文整合）
- [x] `tests/mcp/test_server.py`: patch ツールの `on_conflict` 既定値伝播を検証するテストを追加/更新する
- [x] `tests/mcp/test_patch_runner.py`: 未作成 `out_dir` 保存成功ケースを追加する
- [x] `tests/mcp/test_patch_runner.py`: `.xls` + COM不可エラーの期待値（型・文言）を追加する
- [x] `ruff check .` / `mypy src/exstruct --strict` / 関連pytest を実行し、結果を記録する
