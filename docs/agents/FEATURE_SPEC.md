# Feature Spec for AI Agent (Phase-by-Phase)

## Feature Name

MCP UX Hardening Phase 2 (Claude Review Closure)

## 背景

`review.md` のレビュー結果を一次情報として採用し、MCP UX の未解決課題を解消する。

既に改善済みの項目:

1. 色概念の分離（`color` / `fill_color`）
2. `draw_grid_border` の `range` shorthand
3. `out_path` の root 診断改善

今回の対象は、上記以外の未解決課題（書式一括化、テーブルスタイル、検証UX、成果物連携、大量操作、入力スキーマ可視化）とする。

## 目的

1. 書式操作時の試行錯誤回数を削減する
2. 大量操作時の失敗コストを削減する
3. 生成ファイル受け渡しをチャット内ワークフローで完結させる
4. 実行前に `op` 単位の正しい入力を確認できるようにする

## 非目的

1. 既存 `set_bold` / `set_font_size` / `set_fill_color` などの削除
2. 既存 `exstruct_patch` の原子性（all-or-nothing）の撤廃
3. MCP外部プロダクト固有実装（Claude専用SDK実装）

## スコープ

### FS-01: Validation UX強化

既知の誤入力に対する正規化と、自己修復しやすいエラー情報を追加する。

1. alias 正規化
   1. `set_alignment.horizontal -> horizontal_align`
   2. `set_alignment.vertical -> vertical_align`
   3. `set_fill_color.color -> fill_color`
2. `PatchErrorDetail` 拡張
   1. `hint: str | None`
   2. `expected_fields: list[str]`
   3. `example_op: str | None`
3. エラー文面方針
   1. 何が違うか
   2. 正しい引数名
   3. 最小JSON例

### FS-02: `set_style` 追加

単一opで複数書式を適用できるようにする。

1. 新規 `PatchOpType`: `set_style`
2. 対象指定
   1. `cell` または `range` のどちらか一方（exactly one）
3. 指定可能属性
   1. `bold`
   2. `font_size`
   3. `color`
   4. `fill_color`
   5. `horizontal_align`
   6. `vertical_align`
   7. `wrap_text`
4. 少なくとも1属性必須
5. 既存上限 `_MAX_STYLE_TARGET_CELLS` を適用
6. 既存個別opは後方互換維持

### FS-03: `apply_table_style` 追加

Excelテーブルスタイルを1opで適用できるようにする。

1. 新規 `PatchOpType`: `apply_table_style`
2. 必須
   1. `sheet`
   2. `range`
   3. `style`
3. オプション
   1. `table_name`（未指定時は自動採番）
4. 既存テーブル重複/交差は明示エラー
5. Backend方針
   1. Phase 2 は `openpyxl` 正式対応
   2. `com` は warning を返して `openpyxl` フォールバック

### FS-04: 成果物ミラー（`present_files` 連携）

生成成果物を bridge 先へミラーし、チャット側への受け渡しを容易にする。

1. サーバー起動引数に `--artifact-bridge-dir` を追加
2. `exstruct_make` / `exstruct_patch` 入力に `mirror_artifact: bool = false` を追加
3. 成功時のみ、`mirror_artifact=true` かつ bridge 設定ありでコピー
4. 出力モデルに `mirrored_out_path: str | None` を追加
5. ミラー失敗は処理失敗にせず `warnings` に記録

### FS-05: 大量操作向け分割実行API（新規ツール）

原子性を維持したまま大量opを扱うため、計画と適用を分離する。

1. `exstruct_patch_plan`
   1. 入力: `xlsx_path`, `ops`, `chunk_by`, `max_ops_per_chunk`
   2. 出力: `plan_id`, `chunk_summaries`, `total_ops`
2. `exstruct_patch_apply_chunks`
   1. 入力: `plan_id`, `out_dir`, `out_name`, `backend`
   2. 実行: 内部ステージングで全chunk適用後に最終保存（全体原子性維持）
   3. 失敗時: 最終成果物は生成しない

### FS-06: 入力スキーマ可視化（優先: ツール定義拡充、補完: 確認ツール）

実行前に `op` 単位の入力仕様を確認できるようにする。

1. 優先実装: `exstruct_patch` ツール定義内スキーマ拡充
   1. `op` ごとの required/optional フィールドを明記
   2. フィールド型・制約（exactly one, >0, hex形式など）を明記
   3. `op` ごとの最小実行例を明記
   4. alias（例: `horizontal -> horizontal_align`）の対応を明記
2. 補完実装: スキーマ確認ツール
   1. `exstruct_list_ops`（`op` 一覧と短い説明）
   2. `exstruct_describe_op`（required/optional/constraints/example/aliases）
3. ドリフト防止
   1. ツール定義文言と `describe_op` 生成元は同一メタデータを参照する

### FS-07: シート指定冗長性の削減（top-level `sheet`）

`ops` ごとに `sheet` を重複指定しなくても大量操作を記述できるようにする。

1. `exstruct_patch` / `exstruct_make` 入力に top-level `sheet: str | None = None` を追加
2. 適用ルール
   1. `op.sheet` がある場合は `op.sheet` を優先
   2. `op.sheet` がない場合は top-level `sheet` を補完して使用
3. `add_sheet` は `op.sheet`（または既存 alias `name`）を必須とし、top-level `sheet` は補完しない
4. 非 `add_sheet` 系で `op.sheet` と top-level `sheet` の両方が未指定なら、自己修復可能な明示エラーを返す
5. 後方互換
   1. 既存の `op.sheet` 指定ペイロードは挙動変更なし
   2. mixed運用（同一リクエスト内で一部 `op.sheet` 明示）を許容
6. スキーマ可視化
   1. `exstruct_patch` ミニスキーマを更新し、`sheet` の解決規則を明記
   2. `exstruct_describe_op` の required/optional 表示も top-level `sheet` を反映する

## 主要な公開I/F変更

1. `PatchOpType` 追加
   1. `set_style`
   2. `apply_table_style`
2. `PatchOp` フィールド追加
   1. `style: str | None`（`apply_table_style` 用）
   2. `table_name: str | None`（`apply_table_style` 用）
3. `PatchErrorDetail` 追加フィールド
   1. `hint`
   2. `expected_fields`
   3. `example_op`
4. MCPツール追加
   1. `exstruct_patch_plan`
   2. `exstruct_patch_apply_chunks`
   3. `exstruct_list_ops`
   4. `exstruct_describe_op`
5. MCPツール入出力拡張
   1. `mirror_artifact`（make/patch input）
   2. `mirrored_out_path`（make/patch output）
6. サーバーCLI拡張
   1. `--artifact-bridge-dir`
7. ツール定義拡張
   1. `exstruct_patch` の docstring/description に `op` 別ミニスキーマを追加
8. MCPツール入出力拡張
   1. `sheet: str | None`（patch/make input の top-level デフォルトシート）
9. `PatchOp` のシート解決仕様
   1. `sheet` は「明示時に優先」フィールドとして扱い、最終適用前に解決される

## 受け入れ条件（Acceptance Criteria）

### AC-01 Validation UX

1. 誤引数入力時に `hint` が返る
2. 誤引数入力時に `expected_fields` が返る
3. 誤引数入力時に `example_op` が返る

### AC-02 `set_style`

1. 1op でヘッダ装飾（太字/文字色/背景色/整列）が適用できる
2. `cell`/`range` の同時指定はエラー
3. 属性未指定はエラー

### AC-03 `apply_table_style`

1. 指定範囲にテーブルスタイルが適用される
2. 既存テーブルと交差する範囲指定は明示エラー

### AC-04 成果物ミラー

1. `mirror_artifact=true` かつ bridge 設定ありで `mirrored_out_path` が返る
2. bridge 未設定時は通常処理を継続し warning を返す
3. コピー失敗時も patch/make 結果は成功扱いで warning を返す

### AC-06 入力スキーマ可視化

1. `exstruct_patch` ツール定義だけで主要 `op` の required/optional/example を確認できる
2. `exstruct_list_ops` が利用可能 `op` 一覧を返す
3. `exstruct_describe_op` が `required` / `optional` / `constraints` / `example` / `aliases` を返す

### AC-07 後方互換

1. 既存opの既存入力は挙動変更なし
2. 既存テストが回帰しない

### AC-08 top-level `sheet` 補完

1. top-level `sheet` だけを指定した大量 `ops` で正常適用できる
2. `op.sheet` がある操作は top-level `sheet` より優先される
3. `add_sheet` は `op.sheet`（または `name`）未指定時に明示エラーになる
4. 非 `add_sheet` でシート未解決時は、補完方法を含むエラー情報を返す

## テストケース

1. パラメータ誤り時のヒント返却（`color` vs `fill_color`、`horizontal` vs `horizontal_align`）
2. `set_style` の単セル/範囲/属性未指定エラー
3. `apply_table_style` の正常系/重複テーブルエラー
4. `mirror_artifact` の正常コピー/bridge未設定/コピー失敗warning
7. `exstruct_list_ops` の一覧妥当性
8. `exstruct_describe_op` の required/optional/example 妥当性
9. `exstruct_patch` ツール定義に `op` 別スキーマ情報が含まれること
10. top-level `sheet` 指定時の `op.sheet` 補完と優先順位
11. `add_sheet` の `op.sheet` 必須維持
12. シート未解決時のエラーヒント妥当性

## 前提・デフォルト

1. 既存の `exstruct_patch` 原子性は維持する
2. 新機能は後方互換優先（既存入力/既存レスポンス項目は破壊しない）
3. `mirror_artifact` の既定値は `false`
4. `--artifact-bridge-dir` 未指定時はミラー機能を無効化
5. `apply_table_style` は Phase 2 で openpyxl 優先対応とする
6. 入力スキーマ改善は「ツール定義拡充」を先行し、確認ツールは補完として追加する
7. top-level `sheet` の既定値は `None`（未指定）
8. `add_sheet` は明示的な `op.sheet` 指定を維持する

## 影響範囲

1. `src/exstruct/mcp/server.py`
2. `src/exstruct/mcp/tools.py`
3. `src/exstruct/mcp/patch_runner.py`
4. `src/exstruct/mcp/io.py`（必要時）
5. `docs/mcp.md`
6. `tests/mcp/*`
7. `docs/agents/TASKS.md`
