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

## 次フェーズタスク（便利機能 1〜5）

- [x] `mcp/patch_runner.py`: `PatchRequest` に `dry_run` / `return_inverse_ops` / `preflight_formula_check` を追加する
- [x] `mcp/patch_runner.py`: `PatchResult` に `inverse_ops` / `formula_issues` を追加する
- [x] `mcp/patch_runner.py`: `PatchOp` に `set_range_values` / `fill_formula` / `set_value_if` / `set_formula_if` を追加する
- [x] `mcp/patch_runner.py`: A1範囲パース・範囲サイズ検証・値行列検証ユーティリティを実装する
- [x] `mcp/patch_runner.py`: `set_range_values` の openpyxl 実装を追加する
- [x] `mcp/patch_runner.py`: `fill_formula` の相対参照展開ロジックを実装する（単一行/列制約）
- [x] `mcp/patch_runner.py`: `set_value_if` / `set_formula_if` の比較・`skipped` 差分出力を実装する
- [x] `mcp/patch_runner.py`: 逆パッチ生成ロジック（`applied` のみ、逆順）を実装する
- [x] `mcp/patch_runner.py`: `dry_run=true` で非保存実行に分岐する
- [x] `mcp/patch_runner.py`: 数式ヘルスチェック（`#REF!` 等）を実装し、`formula_issues` に反映する
- [x] `mcp/patch_runner.py`: `preflight_formula_check=true` かつ `level=error` の保存抑止ルールを実装する
- [x] `mcp/tools.py`: `PatchToolInput/Output` の新規フィールドをツールI/Fへ公開する
- [x] `mcp/server.py`: `exstruct_patch` ツール引数に新規3フラグを追加し、ハンドラへ連携する
- [x] `docs/agents/FEATURE_SPEC.md`: 実装差分（命名、エラーコード、制約）を最終反映する
- [x] `tests/mcp/test_patch_runner.py`: `dry_run` 非保存・逆パッチ・範囲操作・条件付き更新・数式検査の単体テストを追加する
- [x] `tests/mcp/test_tools_handlers.py`: 入出力モデル拡張の受け渡しテストを追加する
- [x] `tests/mcp/test_server.py`: 新規ツール引数の受け渡しテストを追加する
- [x] `ruff check .` / `mypy src/exstruct --strict` / 関連pytest を実行し、結果を記録する

## ABC列名キー出力オプション（`alpha_col`）

- [ ] `models/__init__.py`: `col_index_to_alpha()` 変換ユーティリティを追加する
- [ ] `models/__init__.py`: `convert_row_keys_to_alpha()` / `convert_sheet_keys_to_alpha()` を追加する
- [ ] `engine.py`: `StructOptions` に `alpha_col: bool = False` を追加する
- [ ] `engine.py`: `ExStructEngine.extract()` で `alpha_col=True` 時に変換を適用する
- [ ] `__init__.py`: `extract()` / `process_excel()` に `alpha_col` パラメータを追加する
- [ ] `cli/main.py`: `--alpha-col` CLI フラグを追加する
- [ ] `mcp/extract_runner.py`: `ExtractOptions` に `alpha_col` を追加し、`process_excel()` へ連携する
- [ ] `mcp/server.py`: `exstruct_extract` ツールに `alpha_col` パラメータを追加する
- [ ] `tests/models/`: `col_index_to_alpha()` の単体テストを追加する
- [ ] `tests/models/`: `convert_row_keys_to_alpha()` / `convert_sheet_keys_to_alpha()` のテストを追加する
- [ ] `tests/engine/`: `alpha_col=True` でのエンジン抽出テストを追加する
- [ ] `tests/cli/`: `--alpha-col` CLI フラグのテストを追加する
- [ ] `tests/mcp/`: MCP extract の `alpha_col` 連携テストを追加する
- [ ] `ruff check .` / `mypy src/exstruct --strict` / 関連pytest を実行し、結果を記録する
