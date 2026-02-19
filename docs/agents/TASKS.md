# Task List

未完了 [ ], 完了 [x]

## Phase 0: Spec固定
- [x] `exstruct_make` を新規作成専用ツールとして定義する
- [x] `out_path` 必須、`ops` 任意（空許容）を確定する
- [x] 対応拡張子を `.xlsx/.xlsm/.xls` とし、`.xls` はCOM必須で確定する
- [x] 初期シート名 `Sheet1` 正規化を確定する
- [x] `dry_run` 等の拡張フラグを `exstruct_patch` 同等で維持する方針を確定する

## Phase 1: Model / Tool I/F
- [ ] `patch_runner.py` に `MakeRequest` を追加する
- [ ] `patch_runner.py` に `run_make` を追加する
- [ ] `tools.py` に `MakeToolInput` / `MakeToolOutput` を追加する
- [ ] `tools.py` に `run_make_tool` を追加する
- [ ] `__init__.py` の公開シンボルに `make` 系を追加する

## Phase 2: Runner実装
- [ ] `out_path` の拡張子・PathPolicy検証を実装する
- [ ] 空Workbook seed作成（`.xlsx/.xlsm` は openpyxl）を実装する
- [ ] 空Workbook seed作成（`.xls` は COM）を実装する
- [ ] seed初期シートを `Sheet1` に正規化する
- [ ] `run_patch` 再利用で `ops` 適用を実装する
- [ ] 一時seedファイルの確実なクリーンアップを実装する

## Phase 3: Server公開
- [ ] `server.py` に `exstruct_make` ツール登録を追加する
- [ ] `exstruct_make` docstring を追加する
- [ ] `ops` の object/JSON文字列正規化を `patch` と同等に適用する
- [ ] `default_on_conflict` 伝播を `make` にも適用する

## Phase 4: バリデーション / エラーハンドリング
- [ ] `.xls` + `backend=openpyxl` を拒否する
- [ ] `.xls` + COMなし を拒否する
- [ ] `.xls` + `dry_run/return_inverse_ops/preflight_formula_check` を拒否する
- [ ] エラーメッセージを既存 `patch` 系の表現に揃える

## Phase 5: テスト
- [ ] `tests/mcp/test_make_runner.py` を追加する
- [ ] 空作成（opsなし）正常系テストを追加する
- [ ] ops適用正常系テストを追加する
- [ ] `on_conflict` 挙動テストを追加する
- [ ] `.xls` 制約テスト（COM依存・backend制約・拡張フラグ制約）を追加する
- [ ] PathPolicy制約テストを追加する
- [ ] JSON文字列ops受理テストを追加する
- [ ] `tests/mcp/test_server.py` にツール登録・既定値伝播テストを追加する
- [ ] `tests/mcp/test_tools_handlers.py` / `test_tool_models.py` を更新する
- [ ] 既存 `exstruct_patch` 回帰がないことを確認する

## Phase 6: ドキュメント
- [ ] `docs/mcp.md` に `exstruct_make` を追加する
- [ ] `README.md` / `README.ja.md` / `docs/README.en.md` / `docs/README.ja.md` を更新する
- [ ] 必要に応じて `CHANGELOG.md` を更新する
- [ ] `FEATURE_SPEC.md` と本タスクリストの同期を確認する

## Phase 7: 検証
- [ ] `uv run pytest tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_tool_models.py tests/mcp/test_server.py` を実行する
- [ ] `uv run pytest tests/mcp` を実行する
- [ ] `uv run task precommit-run` を実行する
- [ ] 全通過を確認する

## テスト/受け入れ条件
1. `exstruct_make` で新規ブック作成ができる
2. `ops` 適用結果が `patch_diff` に反映される
3. `.xls` 制約エラーが仕様通りに発生する
4. `on_conflict` 挙動が仕様通りに動作する
5. 既存 `exstruct_patch` に回帰がない

## 明示的な前提
1. 新規作成起点は空Workbookのみとする
2. `out_path` は必須とする
3. `ops` は任意（空許容）とする
4. `.xls` はCOM必須で対応する
5. 初期シート名は `Sheet1` に正規化する
