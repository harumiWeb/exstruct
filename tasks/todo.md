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

## Plan (Chart/Table Reliability Improvements 2026-02-27)

- [x] `PatchOp.data_range` を `str | list[str]` へ拡張
- [x] `create_chart` に `chart_title` / `x_axis_title` / `y_axis_title` を追加
- [x] シート名付き `data_range` / `category_range` のバリデーションを追加
- [x] `apply_table_style` の COM 実装を追加
- [x] `service.py` の `apply_table_style -> openpyxl` 強制切替を廃止
- [x] `PatchErrorDetail` に `error_code` / `failed_field` / `raw_com_message` を追加
- [x] COM 実行時の op 単位例外ラップを `Exception` 全般へ拡張
- [x] `op_schema` / `docs/mcp.md` / `README.md` / `README.ja.md` を更新
- [x] `uv run pytest tests/mcp/test_tool_models.py tests/mcp/patch/test_models_internal_coverage.py tests/mcp/patch/test_service.py tests/mcp/test_tools_handlers.py tests/mcp/test_patch_runner.py`
- [x] `uv run task precommit-run`

## Review (Chart/Table Reliability Improvements 2026-02-27)

- Summary:
- `apply_table_style` を COM で実行できるようにし、`backend=auto/com` での openpyxl 強制フォールバックを廃止。`create_chart` は `data_range: list[str]` とシート名付き範囲、`chart_title`/`x_axis_title`/`y_axis_title` に対応。
- Error UX:
- COM 実行時の op 例外ラップを `Exception` 全般に拡張し、`PatchErrorDetail` へ `error_code` / `failed_field` / `raw_com_message` を追加。
- Verification:
- `uv run pytest tests/mcp/test_tool_models.py tests/mcp/patch/test_models_internal_coverage.py tests/mcp/patch/test_service.py tests/mcp/test_tools_handlers.py tests/mcp/test_patch_runner.py` (179 passed)
- `uv run task precommit-run` (ruff / ruff-format / mypy passed)

## Plan (COM Hardening Follow-up 2026-02-27)

- [x] `tasks/feature_spec.md` の Phase 2.1 範囲に沿って詳細仕様を確定
- [x] `apply_table_style` COM Add 呼び出しの互換パターンを実機ログ基準で見直し
- [x] `apply_table_style` の COM 例外を `error_code` 分類へ追加
- [x] Windows + Excel 実機で `apply_table_style` スモークケースを実行
- [x] `docs/mcp.md` / `README.md` / `README.ja.md` に COM前提条件と失敗時対処を追記
- [x] MCP向けの最小サンプル（テーブル作成 + スタイル適用）を docs に追加

## Review (COM Hardening Follow-up 2026-02-27)

- Summary:
- `apply_table_style` の COM 経路で `ListObjects` 取得互換（property/callable）を追加し、`ListObjects.Add` は複数シグネチャ + source文字列フォールバックで再試行するように改善した。
- `ListObjects` の既存テーブル範囲検出は `Address` 正規化を強化し、外部参照付き表記（例: `='[Book]Sheet'!$A$1:$B$2`）でも交差判定できるようにした。
- Error UX:
- `apply_table_style` 向けに `table_style_invalid` / `list_object_add_failed` / `com_api_missing` を分類対象に追加し、該当ケースの `hint` を追加した。
- Verification:
- `uv run pytest tests/mcp/patch/test_models_internal_coverage.py tests/mcp/patch/test_service.py` (56 passed)
- `uv run task precommit-run` (ruff / ruff-format / mypy passed)
- `uv run python -c \"... run_make(... backend='com' ... apply_table_style ...)\"` で `engine='com'` / `error=None` を確認
- Artifact:
- `c:\\dev\\Python\\exstruct\\drafts\\apply_table_style_smoke_local_20260227.xlsx`

## Plan (Review Fix: COM fallback + failed_field 2026-02-27)

- [x] Analyze reviewer findings and locate root causes in `service.py` and `internal.py`.
- [x] Implement `backend=auto` fallback path for COM-originated `PatchOpError`.
- [x] Fix `sheet not found` failed-field classification to support `category_range` context.
- [x] Add regression tests in `tests/mcp/patch/test_service.py` and `tests/mcp/patch/test_models_internal_coverage.py`.
- [x] Run targeted pytest and verify green.

## Review (Review Fix: COM fallback + failed_field 2026-02-27)

- Summary:
- Restored `backend=auto` recovery by allowing fallback when COM op failures are wrapped in `PatchOpError` but still carry COM-runtime markers.
- Corrected `sheet not found` field mapping to classify `category_range` when the message indicates category context.
- Added focused regression tests for both behaviors.
- Verification:
- `uv run pytest tests/mcp/patch/test_service.py tests/mcp/patch/test_models_internal_coverage.py` (51 passed)
- `uv run task precommit-run` (ruff / ruff-format / mypy passed)
- Risks:
- Message-based classification still depends on stable wording; future message format changes can impact field inference.

## Plan (Review Fix: ListObjects property accessor 2026-02-27)

- [x] Reproduce reviewer finding in `internal.py` and confirm pre-resolution `callable` guard blocks property accessor compatibility.
- [x] Remove early `callable` guard and delegate ListObjects resolution to `_resolve_xlwings_list_objects`.
- [x] Add regression test to verify `_apply_xlwings_apply_table_style` succeeds when `ListObjects` is a property collection.
- [x] Run targeted pytest and `uv run task precommit-run`.

## Review (Review Fix: ListObjects property accessor 2026-02-27)

- Summary:
- Fixed `apply_table_style` COM path to support both callable/property `ListObjects` by removing the redundant early callable check.
- Added regression test covering property-style `ListObjects` to prevent future regressions at call-site level.
- Verification:
- `uv run pytest tests/mcp/patch/test_models_internal_coverage.py -k "apply_table_style_accepts_property_list_objects or resolve_xlwings_list_objects_uses_collection_like_accessor"` (2 passed, 46 deselected)
- `uv run task precommit-run` (ruff / ruff-format / mypy passed)

## Plan (MCP usability follow-up from Claude review 2026-02-27)

- [ ] `tasks/feature_spec.md` の新規specに沿って実装順（P0/P1）を確定
- [ ] P0: `_patched` 連鎖を止める出力名ポリシーを実装
- [ ] P0: 出力名ポリシー変更の回帰テストを追加（`tests/mcp/shared/test_output_path.py` ほか）
- [ ] P0: Claude連携向け `--artifact-bridge-dir` / `mirror_artifact` の利用ガイドを `docs/mcp.md` に追加
- [ ] P1: `create_chart` + `apply_table_style` 同時指定エラーに制約理由を追加
- [ ] P1: 同時指定時エラーメッセージの回帰テストを追加（`tests/mcp/patch/test_service.py`）
- [ ] `uv run pytest`（対象テスト）を実行
- [ ] `uv run task precommit-run` を実行

## Review (MCP usability follow-up from Claude review 2026-02-27)

- Summary:
  - (pending)
- Verification:
  - (pending)
- Risks:
  - (pending)
