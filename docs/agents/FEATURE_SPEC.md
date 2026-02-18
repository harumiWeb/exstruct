# Feature Spec for AI Agent (Phase-by-Phase)

## 1. Feature
Issue: MCP patch backend selection + COM-first execution for existing ops

## 2. Goal
既存の `exstruct_patch` 編集機能を、実行環境に応じて `COM(xlwings)` と `openpyxl` で安全に切替できるようにする。

## 3. In Scope
- `exstruct_patch` 入力に `backend` を追加
- `PatchResult` / MCP出力に実行バックエンド `engine` を追加
- 既存opを COM 経路でも実行可能に拡張
- COM失敗時のフォールバック規約を明確化
- 既存 `patch_diff` 構造との互換性維持

## 4. Out of Scope
- 新規の図形編集op / グラフ編集opの追加
- `mcp` ディレクトリの再編（サブフォルダ化）
- 既存op名や出力スキーマの破壊的変更

## 5. Public Interface Changes

### 5.1 `PatchRequest` / `PatchToolInput`
- `backend: Literal["auto", "com", "openpyxl"] = "auto"` を追加

### 5.2 `PatchResult` / `PatchToolOutput`
- `engine: Literal["com", "openpyxl"]` を追加
- 実際に使用したバックエンドを返却

### 5.3 `server.py` (`exstruct_patch` docstring)
- `backend` 引数の仕様（既定値、選択ルール、制約）を明記

## 6. Backend Policy

### 6.1 `backend="auto"`
- COM利用可能なら `com` を優先
- COM利用不可なら `openpyxl`
- ただし以下が `True` の場合は `openpyxl` を選択（現行互換）
  - `dry_run`
  - `return_inverse_ops`
  - `preflight_formula_check`

### 6.2 `backend="com"`
- COM利用不可なら明示エラー
- `.xls` は COM 経路で処理可
- `dry_run` / `return_inverse_ops` / `preflight_formula_check` 併用時は明示エラー

### 6.3 `backend="openpyxl"`
- `.xls` は明示エラー
- `.xlsx` / `.xlsm` のみ処理

### 6.4 COM実行失敗時
- `backend="auto"` かつ入力が `.xlsx` / `.xlsm` の場合のみ `openpyxl` にフォールバックし warning 返却
- `backend="com"` はフォールバックせずエラー返却

## 7. Existing Op Coverage on COM Path
今回の拡張対象（既存op）:
- `set_value`
- `set_formula`
- `add_sheet`
- `set_range_values`
- `fill_formula`
- `set_value_if`
- `set_formula_if`
- `draw_grid_border`
- `set_bold`
- `set_fill_color`
- `set_dimensions`
- `merge_cells`
- `unmerge_cells`
- `set_alignment`

補足:
- `restore_design_snapshot` は内部用途のため当面 openpyxl 専用とする

## 8. Validation Rules
- `backend="com"` と `dry_run/return_inverse_ops/preflight_formula_check` の同時指定を禁止
- `backend="openpyxl"` + `.xls` はエラー
- `backend="com"` + COM不可環境はエラー
- 既存opごとの入力バリデーションは維持

## 9. Test Scenarios
- backend指定ごとの経路選択
- `engine` 返却値検証
- COM不可時のエラー検証
- COM失敗時の `auto` フォールバック検証
- 既存opのCOM経路適用結果検証
- `patch_diff` 構造互換性検証
- 既存warning挙動互換性検証

## 10. Acceptance Criteria
- `backend=auto` で COM 優先動作
- 非COM環境で openpyxl へ安全に切替
- `backend=com` の制約違反で明示エラー
- `engine` が正しく返却される
- `uv run pytest tests/mcp` が通過
- `uv run task precommit-run` が通過

## 11. Assumptions / Defaults
- 図形/グラフ編集opの追加は今回は行わない
- `mcp` サブフォルダ再編は今回は行わない
- 互換性優先で既存スキーマを維持する
- 更新対象ドキュメントは `docs/agents/FEATURE_SPEC.md` と `docs/agents/TASKS.md`
