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

## Feature Name

MCP Patch chart/table reliability improvements (Phase 2)

## Goal

`exstruct_patch` / `exstruct_make` でのグラフ作成とテーブル化の実運用性を高め、AIエージェントが少ない試行回数で安定してExcel自動生成できるようにする。

## Scope

### In Scope

- `apply_table_style` の COM 実装（openpyxl への強制フォールバック廃止）
- `create_chart.data_range` の `list[str]` 対応（非連続/複数系列）
- `create_chart` のシート名付き範囲（`Sheet!A1:B10`）対応
- `create_chart` の明示タイトル設定
  - `chart_title`, `x_axis_title`, `y_axis_title`
- COM 実行失敗時の `PatchErrorDetail` 拡張
  - `error_code`, `failed_field`, `raw_com_message`
- op 単位での COM 例外ラップ（`op_index` 特定性を向上）
- `op_schema` / `docs` / `README` の仕様同期

### Out of Scope

- `create_chart` と `apply_table_style` の同一リクエスト同時実行
- `create_chart` の `dry_run` 対応
- 既存グラフの編集・削除

## Public API / Type Changes

- `PatchOp.data_range`: `str | None` -> `str | list[str] | None`
- `PatchOp` に以下を追加
  - `chart_title: str | None`
  - `x_axis_title: str | None`
  - `y_axis_title: str | None`
- `PatchErrorDetail` に以下を追加
  - `error_code: str | None`
  - `failed_field: str | None`
  - `raw_com_message: str | None`

## Acceptance Criteria

- `backend=auto/com` で `apply_table_style` が COM 経由で実行される
- `create_chart` が `data_range: list[str]` を受理・実行できる
- `create_chart` がシート名付き範囲を受理・実行できる
- `create_chart` でタイトル/軸タイトルを明示設定できる
- COMエラー時に `PatchErrorDetail.error_code` が設定される

## Feature Name

MCP Patch COM hardening follow-up (Phase 2.1)

## Goal

`apply_table_style` を含む COM 系処理の安定性をさらに高め、Excel バージョン差・入力差による失敗率を下げる。

## Scope

### In Scope

- `apply_table_style` の COM Add 呼び出し互換性を追加検証
- `ListObjects` 既存テーブル検出ロジックの堅牢化（Address形式差分吸収）
- COMエラー分類の追加（table style不正、ListObject生成失敗など）
- `apply_table_style` 専用の実運用サンプル（README / docs）追加
- Windows + Excel 実環境スモーク手順の明文化

### Out of Scope

- `create_chart` と `apply_table_style` の同時リクエスト対応
- openpyxl 側の機能拡張

## Acceptance Criteria

- `apply_table_style` の COM 実行が複数 Excel 環境で再現可能に成功する
- 失敗時に `error_code` と修正ヒントで原因特定が可能
- ドキュメントだけで COM 前提条件と制約が判断できる

## Detailed Spec (Phase 2.1)

### 1) `ListObjects.Add` 互換呼び出し

- `apply_table_style` は COM `ListObjects.Add` を複数シグネチャで順次試行する。
- 試行順:
  - `Add(1, Source)`
  - `Add(1, Source, None, 1)`
  - `Add(1, Source, None, 1, None)`
  - `Add(1, Source, None, 1, None, None)`
  - `Add(SourceType=1, Source=...)`
  - `Add(SourceType=1, Source=..., XlListObjectHasHeaders=1)`
- `Source` は `Range API` と `A1文字列` の両方を候補にする（Excelバージョン差吸収）。
- 全試行失敗時は `ValueError("apply_table_style failed to add table ...")` を返す。

### 2) 既存テーブル範囲検出の正規化

- COM `Address` 取得は複数シグネチャでフォールバックし、`$` / `=` / シート修飾を除去して比較する。
- 例:
  - `='[Book1.xlsx]Sales Data'!$B$2:$D$11` -> `B2:D11`
  - `Sheet1!$A$1:$C$9` -> `A1:C9`
- 正規化後の範囲を重複判定 (`_ranges_overlap`) に利用する。

### 3) 失敗分類 (`PatchErrorDetail.error_code`)

- `apply_table_style` で次を追加:
  - `table_style_invalid` (`failed_field="style"`)
  - `list_object_add_failed` (`failed_field="range"`)
  - `com_api_missing` (`failed_field="range"` or `"style"`)
- `hint` には `style` 名修正や `range` 見直し（ヘッダー含む連続範囲）を案内する。

### 4) Windows + Excel スモーク検証

- ツール: `exstruct_make`
- 条件: `backend="com"`、`set_range_values` -> `apply_table_style` の2 op。
- 成功判定:
  - 出力 `.xlsx` が生成される
  - `patch_diff` の `apply_table_style` が `applied`
  - `error` が `null`

## Feature Name

Patch service fallback resilience and failed_field precision (Review Fix 2026-02-27)

## Goal

- Preserve `backend=auto` resilience by keeping COM runtime-error fallback to openpyxl.
- Improve error diagnosis precision by classifying `sheet not found` to the correct range field.

## Scope

### In Scope

- `service.run_patch`: allow openpyxl fallback when COM path raises `PatchOpError` caused by COM runtime errors.
- `_classify_known_patch_error`: detect `category_range` for `sheet not found` messages with category context.
- Add regression tests for both behaviors.

### Out of Scope

- General COM/openpyxl backend policy redesign.
- New error codes or broad message taxonomy changes.

## Acceptance Criteria

- `backend=auto` falls back to openpyxl when COM op-level exception is wrapped as `PatchOpError` with COM-runtime signal.
- `sheet not found` classification reports `failed_field="category_range"` when message context is category range.
- Target tests pass.

## Feature Name

Review Fix: apply_table_style ListObjects property accessor path (2026-02-27)

## Goal

- Ensure `apply_table_style` works when COM `sheet.api.ListObjects` is exposed as a property collection (non-callable).
- Preserve callable accessor compatibility via `_resolve_xlwings_list_objects`.

## Scope

### In Scope

- Remove pre-resolution `callable` guard in `_apply_xlwings_apply_table_style`.
- Route all `ListObjects` resolution through `_resolve_xlwings_list_objects`.
- Add regression test that exercises `apply_table_style` with property-style `ListObjects`.

### Out of Scope

- Changes to COM Add fallback sequence.
- Changes to table style error classification.

## Acceptance Criteria

- `apply_table_style` does not fail early with `sheet ListObjects COM API` when `ListObjects` is a collection property.
- Existing missing-API error behavior is preserved for truly absent `ListObjects`.
- Target regression test passes.

## Feature Name

MCP usability follow-up from Claude Desktop review (2026-02-27)

## Goal

Claude Desktop review で指摘された実運用上の摩擦を優先度順に解消し、
Excel編集体験を壊さずに「出力ファイル運用」「ファイル受け渡し」「操作制約の理解性」を改善する。

## Scope

### In Scope

- P0: `_patched` の連鎖増殖を防ぐ出力名ポリシー改善
  - 既定名が既に `*_patched` のときに再度 `_patched` を重ねない
  - 必要に応じて in-place overwrite を選びやすいオプションを追加
- P0: Claude Desktop 向けファイル受け渡しUXの改善
  - `--artifact-bridge-dir` と `mirror_artifact` の利用導線強化
  - ドキュメントに「Claude連携の推奨起動例」を追加
- P1: `create_chart` + `apply_table_style` 同時指定時のUX改善
  - エラーメッセージに制約理由（backend/engine制約）を明示
  - 将来的な自動分割実行に備えた実装ポイント整理

### Out of Scope

- `--root` 必須方針の撤廃
- セキュリティ境界を弱めるパス制御変更
- `create_chart` と `apply_table_style` の同時実行機能をこのフェーズで完全実装

## Public API / Behavior Changes

- デフォルト出力名生成の振る舞い改善（`*_patched_patched` を抑制）
- `PatchResult` / `MakeResult` の warning 文言を改善し、ユーザーが次アクションを取りやすくする
- ドキュメントに Claude Desktop 連携の設定例を追加

## Acceptance Criteria

- 同じファイルに連続パッチしても、デフォルトで `_patched` が無限に連鎖しない
- `mirror_artifact=true` 利用時の手順が docs のみで再現できる
- `create_chart` + `apply_table_style` の同時指定エラーで「なぜ不可か」が明示される

## Feature Name

MCP Patch COM simultaneous execution for `create_chart` + `apply_table_style` (Phase 3)

## Goal

`create_chart` と `apply_table_style` を同一 `PatchRequest` で実行可能にし、
複数リクエスト分割なしで実運用のExcel生成を完結できるようにする。

## Background / Supersedes

- 本フェーズは、過去フェーズで定義していた
  「`create_chart` と `apply_table_style` の同時リクエスト不可」制約を上書きする。
- ただし `create_chart` が COM専用である方針自体は維持する。

## Scope

### In Scope

- `ops` に `create_chart` と `apply_table_style` が同時に含まれる場合を許可する。
- mixed request は op記述順を保ったまま単一COM実行で処理する。
- `backend=auto` では、mixed request 時に COM が利用可能なら COM を選択する。
- `backend=com` では mixed request を通常ケースとして許可する。
- mixed request の失敗時に `PatchErrorDetail` で原因特定しやすいメッセージを返す。
- `op_schema` / `docs/mcp.md` / `README(.ja).md` の仕様文言を同時更新する。

### Out of Scope

- openpyxl での `create_chart` 実装（`create_chart` のCOM専用は維持）
- mixed request の自動分割実行（1リクエストを内部で2リクエスト化）
- `create_chart` の `dry_run` / `return_inverse_ops` / `preflight_formula_check` 対応

## Behavior Policy

- mixed request (`create_chart` + `apply_table_style`) は **COM専用** とする。
- `backend=openpyxl` で mixed request が渡された場合は入力検証エラーとする。
- `backend=auto` で COM 不可の場合は明示エラーとする（openpyxlフォールバックしない）。
  - 理由: `create_chart` 自体が openpyxl 非対応のため。

## Public API / Type Changes

- `PatchOp` / `PatchRequest` の型追加は不要（既存型のまま対応可能）。
- 変更はバリデーション・実行ポリシーとメッセージの挙動差分のみ。

## Implementation Points

- `src/exstruct/mcp/patch/service.py`
  - mixed-op reject ガードを撤廃し、mixed request を通常処理へ流す。
- `src/exstruct/mcp/patch/models.py` / `src/exstruct/mcp/patch/internal.py`
  - mixed request の backend 制約を validation に明示反映する。
- `src/exstruct/mcp/patch/runtime.py`
  - mixed request 時のエンジン選択・fallback条件を `create_chart` 制約と整合させる。
- `tests/mcp/patch/test_service.py`
  - mixed request 許可ケース（`backend=auto/com`）と COM不可ケースを追加する。
- `tests/mcp/test_patch_runner.py`
  - mixed request の入力制約テスト（`backend=openpyxl` 拒否など）を追加する。

## Tests

- 成功系:
  - `backend=com`: mixed request が `engine="com"` で成功する。
  - `backend=auto` + COM available: mixed request が `engine="com"` で成功する。
- 失敗系:
  - `backend=openpyxl`: mixed request は明示エラーになる。
  - `backend=auto` + COM unavailable: mixed request は明示エラーになる。
- 回帰:
  - `create_chart` 単体の既存制約（COM専用、dry_run不可等）が維持される。
  - `apply_table_style` 単体の openpyxl/COM 既存挙動が破壊されない。

## Acceptance Criteria

- `create_chart` と `apply_table_style` を同一 `PatchRequest` で実行できる。
- mixed request 実行時の `engine` は常に `com` になる。
- mixed request で openpyxl に誤フォールバックしない。
- エラー時に「COM必須」と「なぜ失敗したか」が文言から判別できる。
- 追加テストと `uv run task precommit-run` が通過する。
