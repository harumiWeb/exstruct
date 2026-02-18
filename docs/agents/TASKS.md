# Task List

未完了 [ ], 完了 [x]

## Phase 0: Spec固定
- [x] `PatchOp` の新opと新規フィールド定義を確定する
- [x] 各opの必須/禁止フィールド仕様を確定する
- [x] `.xls` 制約とopenpyxl強制方針を確定する

## Phase 1: Model/Validation実装
- [x] `PatchOpType` に新opを追加する
- [x] `PatchOp` にデザイン編集用フィールドを追加する
- [x] 新op用バリデーション関数を追加する
- [x] 列指定正規化（記号/数値両対応）ヘルパーを実装する
- [x] 色コード正規化ヘルパーを実装する

## Phase 2: Openpyxl適用ロジック実装
- [x] `draw_grid_border` 適用関数を実装する
- [x] `set_bold` 適用関数を実装する
- [x] `set_fill_color` 適用関数を実装する
- [x] `set_dimensions` 適用関数を実装する
- [x] `_apply_openpyxl_op` に新op分岐を追加する
- [x] `_requires_openpyxl_backend` に新opを追加する

## Phase 3: Inverse Ops対応
- [x] スタイル・寸法のスナップショットモデルを追加する
- [x] 逆操作生成ロジックを実装する
- [x] `restore_design_snapshot` 適用ロジックを実装する
- [x] `return_inverse_ops` で新opの逆操作返却を有効化する

## Phase 4: サーバー/ツールI/F更新
- [x] `server.py` の `exstruct_patch` docstringに新op説明を追加する
- [x] 必要に応じて `tools.py` 型定義説明を更新する
- [x] エラーメッセージをAIが理解しやすい文言に統一する

## Phase 5: テスト追加
- [x] `tests/mcp/test_patch_runner.py` に新op正常系テストを追加する
- [x] `tests/mcp/test_patch_runner.py` に新op異常系テストを追加する
- [x] inverse往復テストを追加する
- [x] `tests/mcp/test_server.py` に新op JSON文字列ops受理テストを追加する
- [x] 必要なら `tests/mcp/test_tool_models.py` を更新する

## Phase 6: ドキュメント更新
- [x] `docs/mcp.md` に新op仕様と利用例を追記する
- [x] `README.md` / `README.ja.md` に機能追加を追記する
- [x] `CHANGELOG.md` と release note を更新する
- [x] `docs/agents/FEATURE_SPEC.md` と `docs/agents/TASKS.md` を最終同期する

## Phase 7: 検証
- [x] `uv run pytest tests/mcp` を実行する
- [x] `uv run task precommit-run` を実行する
- [x] 失敗時は修正して再実行し、全通過を確認する

## テスト/受け入れ条件
1. 回帰なし（既存op全維持）。
2. 新4機能が `exstruct_patch` で一貫利用可能。
3. 静的解析・テストが0エラー。

## 明示的な前提
1. 更新対象は `docs/agents/FEATURE_SPEC.md` と `docs/agents/TASKS.md`。
2. MVPは使いやすさ優先で入力自由度を絞る。
3. 実装着手時にこの定義を基準仕様として扱う。
