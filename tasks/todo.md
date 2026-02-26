## Plan

- [x] `tasks/feature_spec.md` を Phase 1 仕様へ刷新
- [x] `tasks/todo.md` を今回実装計画に刷新
- [x] `create_chart` の chart type 共通定義モジュールを追加（`chart_types.py`）
- [x] `models.py` の `chart_type` validator を 8種 + alias 対応へ更新
- [x] `internal.py` の `chart_type` validator / 実行時解決を共通定義へ統一
- [x] `op_schema.py` の `create_chart` 制約文言を更新
- [x] `docs/mcp.md` / `README.md` / `README.ja.md` の対応種別記載を更新
- [x] テスト追加・更新
- [x] `uv run pytest tests/mcp/test_tool_models.py tests/mcp/patch/test_models_internal_coverage.py tests/mcp/test_tools_handlers.py`
- [x] `uv run task precommit-run`

## Test Cases

- [x] `create_chart` で 8 種別を受理できる
- [x] alias 入力（`column_clustered`, `bar_clustered`, `xy_scatter`, `donut`）を正規化できる
- [x] 未対応 `chart_type` で明示エラーになる
- [x] `describe_op(create_chart)` が拡張後制約を返す
- [x] `_resolve_chart_type_id` が 8 種別 + alias を期待IDへ解決できる

## Review

- Summary:
- `create_chart` の `chart_type` を主要8種へ拡張し、alias正規化とCOM ChartType ID解決を共通モジュール化した。`models.py` / `internal.py` / `op_schema.py` / `docs` / `README` / テストを同期更新し、対象pytestとprecommit（ruff/format/mypy）を通過した。
- Risks:
  - COM実環境差異（Excelバージョン依存）により一部種別の表示差が出る可能性
- Follow-ups:
  - Phase 2: `bubble`, `stock`, `surface`, `combo` 対応
  - Phase 2: `chart_subtype`（stacked/100/markers）設計

## Plan (Runtime Chart Creation Test 2026-02-26)

- [x] `tasks/feature_spec.md` にランタイムテスト仕様（`exstruct_make` 入出力契約）を追記
- [x] `drafts` 配下に 8種グラフ作成用の新規 `.xlsx` 出力パスを確定
- [x] `exstruct_make` でテストデータ投入 + 8種 `create_chart` を実行
- [x] 生成ファイル存在と tool レスポンスを確認
- [x] `Review` に結果を記録

## Review (Runtime Chart Creation Test 2026-02-26)

- Summary:
- `mcp__exstruct__exstruct_make`（`engine: com`）で `Charts` シートにテストデータを投入し、`line/column/bar/area/pie/doughnut/scatter/radar` の 8種を `create_chart` で作成できた。`patch_diff` で全 op が `applied`。
- Artifact:
  - `c:\\dev\\Python\\exstruct\\drafts\\create_chart_all_types_20260226.xlsx`
- Verification:
  - `drafts` 一覧に対象ファイルが存在（22.11 KB）
  - ファイル情報の `created/modified` は 2026-02-26 22:12:45 JST
- Risks:
  - Excel COM 実行結果の見た目差（バージョン・環境依存）は残る
