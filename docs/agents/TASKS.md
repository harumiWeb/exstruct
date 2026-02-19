# Task List

未完了 [ ], 完了 [x]

## Phase 0: Spec固定
- [x] `set_font_size` のI/F仕様（対象・制約）を確定する
- [x] バリデーション方針（`cell`/`range`, `font_size > 0`）を確定する
- [x] openpyxl/COM の実装方針を確定する

## Phase 1: Model/Server I/F
- [x] `PatchOp` の `op` 許可一覧に `set_font_size` を追加する
- [x] `PatchOp` に `font_size` フィールドを追加する
- [x] `set_font_size` 用バリデーション関数を追加する
- [x] `server.py` の `exstruct_patch` docstring に `set_font_size` を追記する

## Phase 2: Patch Runner実装
- [x] openpyxl経路に `set_font_size` 適用処理を追加する
- [x] COM経路に `set_font_size` 適用処理を追加する
- [x] `patch_diff` の出力形式を既存スタイルop互換で実装する
- [x] 既存スタイルop共通処理への統合可否を確認する

## Phase 3: テスト
- [x] `set_font_size` 正常系（cell）テストを追加する
- [x] `set_font_size` 正常系（range）テストを追加する
- [x] `font_size <= 0` 異常系テストを追加する
- [x] `cell`/`range` 同時指定・未指定の異常系テストを追加する
- [x] openpyxlで既存フォント属性保持テストを追加する
- [x] COM経路テスト（実行可能範囲）を追加する
- [x] 既存回帰テストが通ることを確認する

## Phase 4: ドキュメント
- [x] `docs/mcp.md` に `set_font_size` を追記する
- [ ] 必要に応じて `README.md` / `README.ja.md` を更新する
- [ ] `CHANGELOG.md` を更新する
- [x] `docs/agents/FEATURE_SPEC.md` と本タスクリストを同期する

## Phase 5: 検証
- [x] `uv run pytest tests/mcp` を実行する
- [x] `uv run task precommit-run` を実行する
- [x] 全通過を確認する

## テスト/受け入れ条件
1. `set_font_size` で対象セルのサイズが変更される
2. `set_font_size` で対象範囲のサイズが変更される
3. `font_size <= 0` はエラー
4. `cell` と `range` の同時指定はエラー
5. `cell` / `range` 未指定はエラー
6. 既存フォント属性（bold等）が維持される
7. 既存opの回帰がない

## 明示的な前提
1. 今回は `font_size` のみ対象（`font_name`/`font_color` は対象外）。
2. 既存スキーマの破壊的変更は行わない。
3. 更新対象は `docs/agents/FEATURE_SPEC.md` と `docs/agents/TASKS.md` を起点とする。
