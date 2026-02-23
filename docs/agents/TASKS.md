# Task List

未完了 [ ], 完了 [x]

## Epic: MCP UX Hardening Phase 3 (Claude Feedback Triage)

### 0. 仕様固定と設計

- [x] `review.md` の指摘を「今回実装 / 見送り」に再分類
- [x] 公開 I/F 変更一覧（型、ツール定義、レスポンス、ドキュメント）を確定
- [x] 後方互換ポリシー（既存 op 挙動維持）を固定
- [x] 見送り項目（`freeze_panes`、`set_comment`、条件付き書式編集、chunk API）を明文化

完了条件:
- [x] `FEATURE_SPEC.md` とタスクが 1 対 1 で追跡可能

### 1. FS-01 Path UX 改善

- [x] `PathPolicy.ensure_allowed` の root 外エラー文に `exstruct_get_runtime_info` 導線を追加
- [x] 既存の `resolved/root/example_relative` 情報を維持
- [x] テスト追加（導線文言の存在確認）

完了条件:
- [x] パスエラー時に自己修復導線が返る

### 2. FS-02 `exstruct_make` 初期シート挙動改善

- [x] `run_make` に初期シート名解決ロジックを追加
- [x] `sheet` 指定 + 同名 `add_sheet` なしで初期シートを改名
- [x] 同名 `add_sheet` ありの場合は `Sheet1` 維持（後方互換）
- [x] openpyxl seed / COM seed の両経路に適用
- [x] テスト追加（改名ケース、競合回避ケース）

完了条件:
- [x] `make` の `sheet` 指定が直感と整合する

### 3. FS-03 `auto_fit_columns` 実装

- [x] `PatchOpType` に `auto_fit_columns` を追加
- [x] `PatchOp` に `min_width` / `max_width` を追加
- [x] validator 実装（許可フィールド制約、`min<=max`、正値制約）
- [x] openpyxl 実装（使用中列判定 + 文字長推定 + clamp）
- [x] COM 実装（AutoFit + clamp）
- [x] `op_schema` / `describe_op` / patch mini schema へ反映
- [x] テスト追加（全列、混在列指定、境界異常）

完了条件:
- [x] 列幅の自動調整を 1 op で実行できる

### 4. FS-04 大量 `ops` ソフト上限警告

- [x] `ops > 200` の warning を `PatchResult.warnings` に追加
- [x] 実行継続（失敗しない）を実装
- [x] テスト追加（`len(ops)=201`）

完了条件:
- [x] 大量操作時に分割判断のガイドが返る

### 5. FS-05 `set_dimensions` diff 可読性改善

- [x] 行/列ターゲット要約ヘルパーを追加
- [x] `set_dimensions` diff を正規化列ラベル要約で出力
- [x] テスト追加（diff 文言検証）

完了条件:
- [x] 列指定の適用確認が diff で容易になる

### 6. FS-06 ドキュメント更新

- [x] `docs/mcp.md` に `auto_fit_columns` quick guide を追加
- [x] `docs/mcp.md` に `ops` ソフト上限（200）と分割ガイドを追加
- [x] `docs/mcp.md` に in-place 上書き手順を追記
- [x] `docs/mcp.md` に path troubleshooting 導線を追記
- [x] `docs/mcp.md` に `set_dimensions` 列指定の確認ポイントを追記

完了条件:
- [x] Phase 3 仕様と利用ガイドの乖離がない

### 7. 検証・受け入れ

- [x] MCP 関連主要テスト実行
- [x] `uv run task precommit-run` 実行
- [x] AC-01〜AC-06 をチェックリストで確認

完了条件:
- [ ] CI グリーン
- [x] 受け入れ条件を満たす

### 8. FS-07 `make` 競合判定の大文字小文字統一

- [ ] `_resolve_make_initial_sheet_name` の競合判定を case-insensitive に修正
- [ ] `make(sheet="Data")` + `add_sheet("data")` で `Sheet1` 維持となることを実装
- [ ] openpyxl/COM の既存挙動を壊さないことを確認
- [ ] テスト追加（大小文字差分の競合ケース）

完了条件:
- [ ] 大文字小文字差分の `add_sheet` で回帰しない

### 9. FS-08 `auto_fit_columns` openpyxl 1-pass 最適化

- [ ] 列ごとの全シート再走査を廃止し、1 回の走査で列最大長を集計
- [ ] 既存の `columns` / `min_width` / `max_width` の挙動互換を維持
- [ ] 回帰テスト追加（1-pass 実装を担保するテスト）

完了条件:
- [ ] `columns` 省略時でも列数依存の過度な性能劣化が発生しない

## 優先順位

1. P0: 1, 2, 3, 4, 5
2. P1: 6, 8
3. P2: 7, 9

## テストケース（必須追跡）

- [x] `make(sheet="Data")`・同名 `add_sheet` なしで初期シートが `Data`
- [x] `make(sheet="Data")` + `add_sheet("Data")` で競合せず適用
- [x] `auto_fit_columns`（全列 + clamp）
- [x] `auto_fit_columns`（`columns=["A",2]`）
- [x] `auto_fit_columns` の不正境界（`min_width > max_width`）
- [x] `len(ops)=201` の warning と成功継続
- [x] `set_dimensions` diff の列要約表示
- [x] root 外パスエラーで `exstruct_get_runtime_info` 導線を返す
- [ ] `make(sheet="Data")` + `add_sheet("data")` で競合回避できる
- [ ] `auto_fit_columns`（openpyxl）が 1-pass 集計で幅計算する
