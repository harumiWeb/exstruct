# Task List

未完了 [ ], 完了 [x]

## Phase 0: Spec固定
- [ ] `backend` / `engine` のI/F仕様を確定する
- [ ] backend選択ルール（auto/com/openpyxl）を確定する
- [ ] COM失敗時フォールバック条件を確定する

## Phase 1: Model/Server I/F
- [ ] `PatchRequest` に `backend` を追加する
- [ ] `PatchResult` に `engine` を追加する
- [ ] `PatchToolInput` に `backend` を追加する
- [ ] `PatchToolOutput` に `engine` を追加する
- [ ] `server.py` の `exstruct_patch` docstring を更新する

## Phase 2: バックエンドルーティング
- [ ] `run_patch` を `backend` 指定対応に更新する
- [ ] `auto` / `com` / `openpyxl` の分岐を実装する
- [ ] `backend=com` の制約違反エラーを実装する
- [ ] COM失敗時の `auto` 限定フォールバックを実装する
- [ ] warning文言を整理し互換性を確認する

## Phase 3: COM実装拡張（既存op）
- [ ] `set_range_values` / `fill_formula` を COM 経路で対応する
- [ ] `set_value_if` / `set_formula_if` を COM 経路で対応する
- [ ] `draw_grid_border` / `set_bold` / `set_fill_color` を COM 経路で対応する
- [ ] `set_dimensions` / `merge_cells` / `unmerge_cells` / `set_alignment` を COM 経路で対応する
- [ ] `patch_diff` 出力形式が既存互換であることを確認する
- [ ] `restore_design_snapshot` は openpyxl専用のまま維持する

## Phase 4: テスト
- [ ] backend選択ユニットテストを追加する
- [ ] `engine` 返却値のテストを追加する
- [ ] COM不可環境での異常系テストを追加する
- [ ] `backend=com` + `dry_run` 等の矛盾指定テストを追加する
- [ ] COM失敗→openpyxlフォールバック（`auto`のみ）テストを追加する
- [ ] 既存op回帰テストを更新する

## Phase 5: ドキュメント
- [ ] `docs/mcp.md` に `backend` 仕様を追記する
- [ ] `README.md` に patch backend 方針を追記する
- [ ] `README.ja.md` に patch backend 方針を追記する
- [ ] `CHANGELOG.md` を更新する
- [ ] `docs/agents/FEATURE_SPEC.md` と本タスクリストを同期する

## Phase 6: 検証
- [ ] `uv run pytest tests/mcp` を実行する
- [ ] `uv run task precommit-run` を実行する
- [ ] 全通過を確認する

## テスト/受け入れ条件
1. `backend=auto` + COM available -> `engine="com"`
2. `backend=auto` + COM unavailable -> `engine="openpyxl"`
3. `backend=com` + COM unavailable -> エラー
4. `backend=openpyxl` + `.xls` -> エラー
5. `backend=com` + `dry_run=True` -> エラー
6. COM例外注入時、`backend=auto` + `.xlsx` -> openpyxlフォールバック + warning
7. `patch_diff` 構造が既存互換
8. 既存warning挙動の互換性維持

## 明示的な前提
1. 図形/グラフ編集opの追加は今回対象外。
2. `mcp` サブフォルダ再編は今回対象外。
3. `return_inverse_ops` / `dry_run` / `preflight_formula_check` は当面 openpyxl 優先方針を維持する。
4. 更新対象は `docs/agents/FEATURE_SPEC.md` と `docs/agents/TASKS.md` のみ。
