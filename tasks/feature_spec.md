# Feature Spec

## Feature Name

MCP Patch `create_chart` major chart support (Phase 1)

## Goal

`exstruct_patch` / `exstruct_make` の `create_chart` で主要なグラフ種別をより広く扱えるようにし、MVP (`line/column/pie`) から実運用レベルへ拡張する。

## Scope

### In Scope

- `chart_type` の対応拡張（Phase 1）
  - `line`, `column`, `bar`, `area`, `pie`, `doughnut`, `scatter`, `radar`
- `chart_type` 正規化キー方式
  - alias 受理: `column_clustered`, `bar_clustered`, `xy_scatter`, `donut`
- 既存引数仕様の維持
  - `data_range`, `category_range`, `anchor_cell`
  - `width`, `height`, `chart_name`
  - `titles_from_data`, `series_from_rows`
- モデル層と実行層で同一の chart type 対応表を参照（重複定義の解消）
- `op_schema` / `describe_op` / `docs` / `README` 更新

### Out of Scope

- 既存グラフ編集・削除
- グラフ詳細装飾（軸書式、凡例位置、配色指定など）
- openpyxl バックエンドでのグラフ作成
- `bubble/stock/surface/combo` など Phase 2 以降の種別

## Public API / Type Changes

- `PatchOp` のフィールド追加・削除なし（後方互換維持）
- `chart_type` の許可値を以下へ拡張
  - `line`, `column`, `bar`, `area`, `pie`, `doughnut`, `scatter`, `radar`
- alias を受理し、内部で canonical key に正規化

## Validation Rules (`create_chart`)

- 必須: `sheet`, `chart_type`, `data_range`, `anchor_cell`
- 任意: `category_range`, `chart_name`, `width`, `height`, `titles_from_data`, `series_from_rows`
- 制約:
  - `chart_type` は対応キーまたは alias のみ許可
  - `width > 0`, `height > 0`（指定時）
  - `chart_name` は空文字不可
  - A1形式妥当性（`data_range`, `category_range`, `anchor_cell`）
  - 非関連フィールドは拒否
- 未対応値は明示エラー（フォールバックしない）

## Backend Policy

- `create_chart` は COM専用
- `backend=openpyxl` ならエラー
- `backend=auto` で COM不可ならエラー
- `dry_run` / `return_inverse_ops` / `preflight_formula_check` とは併用不可
- `apply_table_style` と同一リクエストで併用不可

## Chart Type Mapping (Phase 1)

- `line` -> `4` (Line)
- `column` -> `51` (ColumnClustered)
- `bar` -> `57` (BarClustered)
- `area` -> `1` (Area)
- `pie` -> `5` (Pie)
- `doughnut` -> `-4120` (Doughnut)
- `scatter` -> `-4169` (XYScatter)
- `radar` -> `-4151` (Radar)

## Implementation Points

- 共通定義: `src/exstruct/mcp/patch/chart_types.py`
- バリデーション: `src/exstruct/mcp/patch/models.py`, `src/exstruct/mcp/patch/internal.py`
- 実行時解決: `src/exstruct/mcp/patch/internal.py`
- スキーマ: `src/exstruct/mcp/op_schema.py`
- ドキュメント: `docs/mcp.md`, `README.md`, `README.ja.md`

## Tests

- `tests/mcp/test_tool_models.py`
  - 新規 chart type 8種の受理
  - alias 正規化
  - 未対応値のエラー
- `tests/mcp/patch/test_models_internal_coverage.py`
  - COM ChartType ID 解決（8種）
  - alias 解決
- `tests/mcp/test_tools_handlers.py`
  - `describe_op(create_chart)` の制約文言更新
- 既存回帰確認
  - `tests/mcp/test_patch_runner.py`
  - `tests/mcp/patch/test_service.py`
  - `tests/mcp/test_server.py`

## Acceptance Criteria

- `create_chart` が 8 種類で利用可能
- MVP の既存入力 (`line/column/pie`) は完全後方互換
- alias 入力が canonical key に正規化される
- 未対応 `chart_type` は明示エラーを返す
- `op_schema` / `describe_op` / ドキュメントが実装と一致する

## Runtime Test Spec (2026-02-26)

### Goal

`exstruct_make` を使って、`create_chart` の 8 種別（`line`, `column`, `bar`, `area`, `pie`, `doughnut`, `scatter`, `radar`）を実際に作成した Excel を `drafts` 配下に新規生成する。

### Tool Call Contract

- Tool: `mcp__exstruct__exstruct_make`
- Input:
  - `out_path: str`
  - `ops: list[PatchOp]`（JSON文字列で渡す）
  - `backend: Literal["auto", "com"]`（本テストでは `auto`）
- Output:
  - `status: str`
  - `output_path: str`
  - `diff: list[object]`
  - `warnings: list[str]`

### Acceptance

- `drafts` に新規 `.xlsx` が生成される
- 8種類すべての `create_chart` オペレーションが成功する
- エラーが返らない
