# Feature Spec

## Feature Name

MCP Patch `create_chart` (新規グラフ作成のみ)

## Goal

`exstruct_patch` / `exstruct_make` の `ops` に `create_chart` を追加し、既存グラフ編集なしで新規作成できるようにする。

## Scope

### In Scope

- `create_chart` op追加
- 対応種別: `line` / `column` / `pie`
- 入力: 範囲指定型（`data_range`, `category_range`, `anchor_cell`）
- 配置: `anchor_cell` 必須、`width`/`height` 任意（points）
- `titles_from_data`（default: true, best-effort）
- `series_from_rows`（default: false）
- `chart_name` 任意（同一sheet内重複はエラー）
- COM専用実行

### Out of Scope

- 既存グラフ編集・削除
- グラフ詳細装飾（軸書式、色詳細、凡例編集など）
- `line/column/pie` 以外の種別

## Public API / Type Changes

- `PatchOpType` に `create_chart` を追加
- `PatchValueKind` に `chart` を追加
- `PatchOp` に以下を追加  
  `chart_type`, `data_range`, `category_range`, `anchor_cell`, `chart_name`, `width`, `height`, `titles_from_data`, `series_from_rows`

## Validation Rules (`create_chart`)

- 必須: `sheet`, `chart_type`, `data_range`, `anchor_cell`
- 任意: `category_range`, `chart_name`, `width`, `height`, `titles_from_data`, `series_from_rows`
- 制約:
  - `chart_type in {"line","column","pie"}`
  - `width > 0`, `height > 0`（指定時）
  - `chart_name` は空文字不可
  - A1形式妥当性（`data_range`, `category_range`, `anchor_cell`）
  - 非関連フィールドは拒否

## Backend Policy

- `create_chart` は COM専用
- `backend=openpyxl` ならエラー
- `backend=auto` で COM不可ならエラー
- `backend=com` は既存COM可用性チェックを利用

## Implementation Points

- `op_schema` / `list_ops` / `describe_op` に `create_chart` を追加
- `internal.py`:
  - openpyxl側は `create_chart` を明示拒否
  - xlwings側に `_apply_xlwings_create_chart` を追加
- `PatchDiffItem.after.kind="chart"` を返却

## Tests

- `tests/mcp/test_tool_models.py`
- `tests/mcp/test_tools_handlers.py`
- `tests/mcp/test_server.py`
- `tests/mcp/patch/test_models_internal_coverage.py`
- `tests/mcp/patch/test_service.py`
- `tests/mcp/test_patch_runner.py`

## Acceptance Criteria

- `create_chart` が `exstruct_patch` / `exstruct_make` で利用可能
- `list_ops` / `describe_op` に表示される
- openpyxl経路では明示エラー
- 既存opに回帰なし
