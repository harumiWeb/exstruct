# Feature Spec for AI Agent (Phase 3)

## Feature Name

MCP UX Hardening Phase 3 (Claude Feedback Triage)

## 背景

`review.md` のフィードバックを精査し、以下を再分類した。

1. 今回実装する項目
2. 今回は見送り、次フェーズ候補に残す項目

Phase 2 までで対応済みの項目（`set_style`、`apply_table_style`、`mirror_artifact`、top-level `sheet`、スキーマ可視化など）は維持し、Phase 3 では初期体験と実運用時の失敗コスト削減に集中する。

## 目的

1. パス解決の初期つまずきを減らす
2. `exstruct_make` の直感的なシート生成挙動を提供する
3. 列幅調整の手動試行錯誤を減らす
4. 大量 `ops` 実行時の判断を支援する
5. `set_dimensions` の実行結果確認を容易にする

## 非目的

1. `freeze_panes` の実装
2. セルコメント追加 (`set_comment`) の実装
3. 条件付き書式編集の実装
4. `apply_table_style` の COM ネイティブ対応
5. `patch_plan` / `patch_apply_chunks` の新規 API 追加
6. 上書き専用の新規 API フラグ追加

## 追加レビュー指摘（未コミット差分レビュー）

未コミット差分レビューで、次の 2 点を修正対象として追加する。

1. `make(sheet=...)` と `add_sheet(...)` の競合判定が大文字小文字差異で漏れる
2. openpyxl の `auto_fit_columns` が列数に比例して全シート走査を繰り返す

## レビュー指摘の採否

| 項目 | 判定 | 方針 |
|---|---|---|
| パス解決が分かりにくい | 実施 | エラーヒント強化 + ドキュメント導線改善 |
| `exstruct_make` の `sheet` 指定で直感的に作成したい | 実施 | 条件付きで初期シート改名 |
| `set_dimensions` の列指定確認がしにくい | 実施 | diff可読性改善 + docs明記 |
| 大量 `ops` の上限が不明 | 実施 | 推奨上限を仕様化し warning 返却 |
| `auto_fit` が欲しい | 実施 | 新規 `auto_fit_columns` op追加 |
| `freeze_panes` が欲しい | 見送り | 次フェーズ候補 |
| 条件付き書式編集 | 見送り | 別フェーズ |
| セルコメント追加 | 見送り | 今回は `auto_fit_columns` 優先 |
| `apply_table_style` の COM warning 解消 | 見送り | COM ネイティブ対応は別フェーズ |
| 上書きモードの分かりにくさ | 部分実施 | API追加せず docs 明確化 |

## スコープ

### FS-01: Path UX 改善

`PathPolicy.ensure_allowed` の root 外エラーに、自己修復導線を追加する。

1. 既存エラーメッセージに以下を追記
   1. `exstruct_get_runtime_info` の利用案内
2. 既存の `resolved=... root=... example_relative=...` 情報は維持

### FS-02: `exstruct_make` の初期シート挙動改善

`MakeRequest.sheet` が指定された場合の初期シート名を改善する。

1. ルール
   1. `sheet` 指定あり かつ 同名 `add_sheet` が `ops` に無い場合
      1. seed workbook の初期シートを `sheet` 名へ改名して開始
   2. 同名 `add_sheet` がある場合
      1. 初期シートは `Sheet1` のまま維持（後方互換）
2. `.xlsx/.xlsm` (openpyxl) と `.xls` (COM seed) の両方で同一ルール

### FS-03: `auto_fit_columns` 追加

列幅の自動調整を 1 op で実行できるようにする。

1. 新規 `PatchOpType`: `auto_fit_columns`
2. 新規 `PatchOp` フィールド
   1. `min_width: float | None`
   2. `max_width: float | None`
3. 入力仕様
   1. 必須: `sheet`
   2. 任意: `columns`, `min_width`, `max_width`
   3. 制約: `min_width > 0`, `max_width > 0`, `min_width <= max_width`
4. 適用対象
   1. `columns` 指定あり: 指定列のみ（`"A"` と `2` の混在可）
   2. `columns` 省略: 使用中列全体
5. バックエンド方針
   1. openpyxl: 文字長ベース推定で幅計算し、必要に応じて clamp
   2. COM: `AutoFit` 実行後に必要に応じて clamp

### FS-04: 大量 `ops` ソフト上限警告

大量実行の判断支援として警告を返す。

1. しきい値: `200`
2. 条件: `len(ops) > 200`
3. 挙動
   1. 処理は継続（失敗にしない）
   2. `PatchResult.warnings` に分割推奨メッセージを追加

### FS-05: `set_dimensions` diff 可読性改善

列指定の適用確認を容易にする。

1. `set_dimensions` の `PatchDiffItem.after.value` を改善
2. 列情報は正規化後の列名を要約表示
   1. 例: `columns=A, B (2)`
3. 行情報も同様に要約表示

### FS-06: ドキュメント改善

`docs/mcp.md` を Phase 3 仕様へ更新する。

1. `auto_fit_columns` の quick guide を追加
2. `set_dimensions` の列指定（文字/数値両対応）を明記
3. `ops` ソフト上限（200）と分割ガイドを追記
4. in-place 上書きの明確な手順を追記
5. パスエラー時の `exstruct_get_runtime_info` 導線を追記

### FS-07: `make` 初期シート競合判定の厳密化

`MakeRequest.sheet` と `add_sheet` の競合判定を Excel のシート名解決規則に合わせる。

1. 競合判定は大文字小文字非依存で行う
   1. `sheet="Data"` と `add_sheet("data")` は競合として扱う
2. 競合時の挙動は既存仕様を維持
   1. 初期シートは `Sheet1` のまま開始する
3. 後方互換
   1. 既存の同一大文字小文字ケース（`Data` と `Data`）の挙動は不変

### FS-08: `auto_fit_columns`（openpyxl）の 1-pass 化

openpyxl バックエンドの列幅推定を 1 回のシート走査で完結させる。

1. 目的
   1. 列ごとの全シート再走査を廃止し、列数増加時の実行時間悪化を抑制する
2. 方針
   1. 1 回の走査で列ごとの最大表示長を集計
   2. 集計結果から対象列の推定幅を算出して clamp を適用
3. 非機能要件
   1. `columns` 省略時でも実行時間のオーダーが `O(セル数 + 対象列数)` に収まること
4. 機能互換
   1. 既存の `min_width` / `max_width` / `columns` の意味は変更しない

## 主要な公開 I/F 変更

1. `PatchOpType` 追加
   1. `auto_fit_columns`
2. `PatchOp` フィールド追加
   1. `min_width: float | None`
   2. `max_width: float | None`
3. `exstruct_patch` の mini schema 追加
   1. `auto_fit_columns` の required / optional / constraints / example
4. `exstruct_describe_op` 対応
   1. `auto_fit_columns` の仕様詳細を返却
5. `PatchResult.warnings` 拡張
   1. `ops > 200` の推奨分割 warning
6. `set_dimensions` diff 表示変更
   1. 列要約の可読化
7. `exstruct_make` seed 作成仕様変更
   1. 条件成立時に初期シート名として `sheet` を採用

## 受け入れ条件（Acceptance Criteria）

### AC-01 Path UX

1. root 外パスエラーに `exstruct_get_runtime_info` 導線が含まれる

### AC-02 `make` 初期シート挙動

1. `make(sheet="Data")` かつ同名 `add_sheet` なしで初期シートが `Data` になる
2. 同名 `add_sheet("Data")` を含む場合は後方互換を維持し、重複エラーを起こさない

### AC-03 `auto_fit_columns`

1. `columns` 省略で使用中列全体が対象になる
2. `columns=["A", 2]` の混在指定が適用される
3. `min_width` / `max_width` で幅が clamp される

### AC-04 大量 `ops` 警告

1. `len(ops)=201` で warning が返る
2. 処理自体は成功する

### AC-05 `set_dimensions` diff

1. 実行結果 diff に正規化済み列ラベル要約が含まれる

### AC-06 後方互換

1. 既存 op の既存入力は挙動変更しない
2. 既存テストが回帰しない

### AC-07 `make` 競合判定（大文字小文字）

1. `make(sheet="Data")` + `add_sheet("data")` で重複エラーを発生させない
2. 上記ケースで初期シートは `Sheet1` を維持し、`add_sheet("data")` が適用される

### AC-08 `auto_fit_columns` 1-pass 性能要件

1. openpyxl 実装で列ごとの全シート再走査を行わない
2. `columns` 省略時の大規模シートでも実行時間が列数に対して過度に悪化しない

## テストケース

1. `make` で `sheet="Data"`・同名 `add_sheet` なしの初期シート名確認
2. `make` で `sheet="Data"` + `add_sheet("Data")` の競合回避確認
3. `auto_fit_columns`（全列対象 + clamp）
4. `auto_fit_columns`（`columns=["A",2]`）
5. `auto_fit_columns` の不正境界（`min_width > max_width`）エラー
6. `len(ops)=201` の warning
7. `set_dimensions` diff の列要約確認
8. root 外パスエラーメッセージの導線確認
9. `make(sheet="Data")` + `add_sheet("data")` で競合回避できること
10. `auto_fit_columns` openpyxl が 1-pass 集計で幅計算すること（回帰防止）

## 前提・デフォルト

1. ソフト上限しきい値は `200`
2. `ops > 200` は warning のみ（失敗にしない）
3. `auto_fit_columns.columns` 未指定時は使用中列全体を対象
4. openpyxl は文字長ベース推定、COM は `AutoFit` ベース
5. `freeze_panes`、`set_comment`、条件付き書式編集は今回スコープ外
6. `patch_plan/apply_chunks` は今回スコープ外
7. 上書き専用 API は追加しない（ドキュメントで運用明確化）
8. Excel シート名競合は大文字小文字非依存として扱う

## 影響範囲

1. `src/exstruct/mcp/io.py`
2. `src/exstruct/mcp/patch_runner.py`
3. `src/exstruct/mcp/op_schema.py`
4. `src/exstruct/mcp/server.py`
5. `docs/mcp.md`
6. `tests/mcp/test_path_policy.py`
7. `tests/mcp/test_make_runner.py`
8. `tests/mcp/test_patch_runner.py`
9. `tests/mcp/test_tool_models.py`
10. `tests/mcp/test_tools_handlers.py`
11. `tests/mcp/test_server.py`
12. `docs/agents/TASKS.md`
