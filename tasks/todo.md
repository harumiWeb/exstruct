# Todo

## 2026-03-06 Backend Metadata Output Follow-up

### Planning

- [x] shape/chart backend metadata の常時出力箇所を調査する
- [x] `include_backend_metadata` を Python API / engine / CLI / MCP の共通オプションとして追加する
- [x] 既定値を metadata 非表示に切り替え、明示指定時のみ出力する
- [x] 直列化 helper / model helper / per-sheet / per-area export の挙動を揃える
- [x] README / API docs / MCP docs / release note / spec を更新する
- [x] 関連 unit test を追加・更新する

### Review

- `serialize_workbook`, `save_sheets`, `save_print_area_views`, `save_auto_page_break_views` に `include_backend_metadata` を追加し、shape/chart metadata を出力時だけ制御するようにした
- `SheetData` / `WorkbookData` / `PrintAreaView` の `to_json` / `to_yaml` / `to_toon` / `save` に同フラグを通し、raw model 直列化の既定値も揃えた
- `FilterOptions.include_backend_metadata`, CLI `--include-backend-metadata`, MCP `options.include_backend_metadata` を追加した
- 追加検証:
  - `uv run pytest tests/models/test_models_export.py tests/engine/test_engine.py tests/cli/test_cli.py tests/mcp/test_extract_runner_utils.py tests/mcp/test_tool_models.py tests/core/test_mode_output.py -k "not mark.com and not standard and not verbose and not line and not connector" -q`
  - `uv run pytest tests/io/test_print_area_views.py tests/core/test_error_handling_exceptions.py tests/export/test_export_requirements.py tests/mcp/test_tools_handlers.py tests/mcp/test_extract_alpha_col.py -q`

## 2026-03-06 README Sync

### Planning

- [x] `README.ja.md` を基準に `README.md` / `docs/README.ja.md` / `docs/README.en.md` の差分を整理する
- [x] `docs/` 配下で必要な画像・リンク・相対パス差分を確認する
- [x] `README.md` を刷新後の日本語版構成に合わせて更新する
- [x] `docs/README.ja.md` を刷新後の日本語版構成に合わせて更新する
- [x] `docs/README.en.md` を刷新後の日本語版構成に合わせて更新する
- [x] 差分確認で 3 ファイルの整合性を確認し、レビュー結果を記録する

### Review

- 3 ファイルとも先頭をロゴ + バッジ構成へ統一し、刷新後 `README.ja.md` に合わせて導入文・章立て・MCP 節を更新した
- `docs/README.ja.md` / `docs/README.en.md` は `assets/...` や `architecture/...` など `docs/` 配下向けの相対パスへ調整した
- `Documentation build` / `MCP追加メモ（UX Hardening）` / 旧 `Migration note` など、基準 README にない節は除去した
- `git diff` と `rg` で対象 3 ファイルに対する主要見出し・MCP 詳細・画像パス・不要節の有無を確認した

## Planning

- [x] issue 56 の本文とコメントを確認し、must-have / nice-to-have / 非ゴールを整理する
- [x] 既存の pipeline / shape / chart / render / MCP 実装を確認し、変更境界を特定する
- [x] `tasks/feature_spec.md` に `libreoffice` mode の仕様、型、fallback 方針を定義する
- [x] 実装順と検証方針をこの `tasks/todo.md` に落とし込む

## Phase 1: Public Contract

- [x] `ExtractionMode` を `light/libreoffice/standard/verbose` に拡張する
- [x] Python API (`extract`, `process_excel`, `StructOptions`, `extract_workbook`) の型と docstring を更新する
- [x] CLI `--mode` choices と help を更新する
- [x] MCP `ExtractRequest`, `server.py`, `docs/mcp.md` の mode 説明を更新する
- [x] `.xls` + `mode="libreoffice"` を早期バリデーションで拒否する

## Phase 2: Pipeline / Runtime

- [x] `resolve_extraction_inputs` に `libreoffice` の既定 include_* を追加する
- [x] `PipelineState` / `FallbackReason` に LibreOffice 用 reason を追加する
- [x] pipeline の rich backend 選択を `light/com/libreoffice` 前提で整理する
- [x] LibreOffice session helper を追加し、headless 起動・一時 profile・timeout・cleanup を実装する
- [x] LibreOffice 不在時の fallback を cells/tables/print_areas/merged_cells 維持で実装する

## Phase 3: Shape / Connector

- [x] OOXML drawing helper を追加し、shape / connector / chart anchor 情報を読めるようにする
- [x] LibreOffice UNO から draw-page shapes を取得する backend を追加する
  - `LibreOfficeSession.extract_draw_page_shapes(...)` と bridge payload を実装する
  - `LibreOfficeRichBackend.extract_shapes(...)` が UNO draw-page 順を canonical source として使う
- [x] LibreOffice bridge payload に draw-page shape と connector direct-ref を追加する
- [x] LibreOffice backend で UNO draw-page payload を shape metadata と connector 解決に統合する
- [x] non-connector shape のみシート内連番 `id` を振る仕様を実装する
- [x] connector 解決を `OOXML explicit ref -> UNO direct ref -> geometry heuristic` の優先順で実装する
  - OOXML connector match が取れない場合でも UNO `StartShape/EndShape` を使って begin/end を復元する
  - direct ref 不可時のみ geometry heuristic に落とす回帰 test を追加する
- [x] `BaseShape` metadata (`provenance`, `approximation_level`, `confidence`) を追加し、COM / LibreOffice 両経路で埋める

## Phase 4: Chart

- [x] OOXML / openpyxl から chart の semantic 情報を抽出する helper を追加する
- [x] LibreOffice UNO から chart geometry 候補を取得する
- [x] OOXML chart と UNO geometry を順序で pairing し、geometry を `Chart` に反映する
- [x] UNO geometry が無い場合は openpyxl anchor を geometry fallback として使う
- [x] `Chart` metadata (`provenance`, `approximation_level`, `confidence`) を追加する

## Phase 5: Verification

- [x] mode validation の unit test を追加する
- [x] `.xls` reject の unit test を追加する
- [x] `sample/flowchart/sample-shape-connector.xlsx` を使った connector graph 回帰 test を追加する
- [x] `sample/basic/sample.xlsx` を使った chart extraction 回帰 test を追加する
- [x] LibreOffice unavailable fallback の unit test を追加する
- [x] 必要なら `pytest.mark.libreoffice` の optional smoke test を追加する
- [x] `uv run pytest` または対象 test を実行して結果を確認する
- [x] `uv run task precommit-run` を実行し、ruff / mypy / format 系の問題が無いことを確認する

## Phase 6: Documentation

- [x] README.md / README.ja.md の mode 説明を更新する
- [x] contributor / architecture / release notes を更新する
- [x] `libreoffice` mode が best-effort であり strict subset ではないことを明記する
- [x] rendering と auto page-break が v1 対象外であることを明記する

## Phase 7: CLI LibreOffice Follow-up

- [x] `tasks/feature_spec.md` に CLI / process API の互換性ルールを追記する
- [x] `engine` / `integrate` に `libreoffice` 非対応オプションの共通バリデーションを追加する
- [x] CLI help を更新し、`--pdf` / `--image` / `--auto-page-breaks-dir` の制約を明記する
- [x] `process_excel(...)` / `ExStructEngine.process(...)` / `extract_workbook(...)` の docstring と例を新仕様に揃える
- [x] CLI / API / extract レイヤーの異常系 test を追加する
- [x] README / CLI Guide / API docs / test requirements を今回の契約に揃える

## Review

- 2026-03-06 draw-page / connector follow-up:
  - LibreOffice bridge に `--kind draw-page` を追加し、`DrawPage` 由来の shape / connector payload を取得可能にした
  - `extract_shapes(mode="libreoffice")` は UNO draw-page 順を canonical source にしつつ、OOXML を type / arrowhead / explicit ref 補完に限定して使う
  - connector 解決順を `OOXML explicit ref -> UNO direct ref -> geometry heuristic` に固定し、UNO-only と explicit-priority の unit test を追加した
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py tests/core/test_mode_output.py -k libreoffice -q` と `RUN_LIBREOFFICE_SMOKE=1` 付き smoke、`uv run task precommit-run` を通した
- 2026-03-06 chart geometry follow-up:
  - LibreOffice 同梱 Python bridge subprocess を追加し、`sheet.getCharts()` + `DrawPage` `OLE2Shape` から chart geometry 候補を取得
  - OOXML chart name / `PersistName` 一致を優先し、残差のみ順序 pairing する `libreoffice` mode の chart geometry 反映を実装
  - `RUN_LIBREOFFICE_SMOKE=1` 付き smoke test で chart geometry が 0 埋めではなく UNO geometry になることを確認

- 2026-03-06 follow-up:
  - `pytest.mark.libreoffice` smoke test と `RUN_LIBREOFFICE_SMOKE=1` gate を追加
  - LibreOffice runtime あり環境で smoke test を実行して通過
- 2026-03-06 CLI libreoffice follow-up:
  - `src/exstruct/constraints.py` を追加し、`mode="libreoffice"` と PDF/PNG rendering / auto page-break export の非対応組み合わせを `ConfigError` で早期拒否する共通ガードを実装
  - `extract_workbook(...)` / `ExStructEngine.process(...)` / CLI help / README / `docs/cli.md` / `docs/api.md` / `docs/agents/TEST_REQUIREMENTS.md` を同じ契約に揃えた
  - `uv run pytest tests/core/test_mode_output.py tests/cli/test_cli.py tests/backends/test_auto_page_breaks.py -q` は `28 passed, 1 skipped` で通過したが、既存 COM テスト由来の Windows COM 例外ログが終了後 stderr に出る現象は残っている
  - `uv run pytest tests/core/test_mode_output.py -k "libreoffice or process_excel_rejects or extract_workbook_rejects" tests/cli/test_cli.py -k "libreoffice" tests/backends/test_auto_page_breaks.py -k "libreoffice or extract_rejects_auto_page_break_flag" -q` は `8 passed`
  - `uv run task precommit-run` は ruff / ruff-format / mypy すべて通過

- 実装状態: 完了
- この時点で完了済み:
  - issue 56 の仕様整理
  - `feature_spec.md` 作成
  - 実装タスク分解
  - public contract / pipeline fallback / OOXML helper / best-effort backend 実装
  - mode / fallback / sample regression / metadata / docs 更新
  - `uv run pytest` 対象群と `uv run task precommit-run` の通過確認
- 実装完了条件:
  - public API / CLI / MCP の mode 追加が揃っている
  - connector graph と chart の best-effort 抽出が動く
  - 既存 COM 挙動に回帰がない
  - test / precommit-run が通る
- 主なリスク:
  - UNO API の環境差
  - connector heuristic の誤接続
  - chart geometry の pairing ずれ

## 2026-03-07 LibreOffice bridge compatibility probe follow-up

### Planning

- [x] `tasks/feature_spec.md` の互換性契約に合わせて、LibreOffice bridge 用 Python 候補の受け入れ条件を `bridge 実行可能` ベースへ更新する
- [x] `src/exstruct/core/_libreoffice_bridge.py` に no-op probe 用の `--probe` を追加し、UNO socket や workbook access を伴わない実行経路を用意する
- [x] `src/exstruct/core/libreoffice.py` の `_python_supports_libreoffice_bridge(...)` を、UNO import だけでなく bundled bridge probe 実行まで検証する実装へ変更する
- [x] `_resolve_python_path(...)` が system Python 候補を選ぶ際、`uno` import は通るが bridge 実行は失敗する候補を reject するようにする
- [x] `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` 指定時も同じ probe で fail-fast し、遅延 `SyntaxError` ではなく明確な incompatible runtime error を返す
- [x] `tests/core/test_libreoffice_backend.py` に、`uno` import success / bridge `SyntaxError` failure の false positive を防ぐ regression test を追加する
- [x] `tests/conftest.py` の LibreOffice runtime smoke gate も、必要なら bridge 実行可能性を反映した判定に揃える
- [x] `README.md` / `README.ja.md` / `docs/agents/TEST_REQUIREMENTS.md` の「compatible system Python」説明を probe ベースの表現へ揃える
- [x] `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py -q` と `uv run task precommit-run` で検証する

### Review

- 受け入れ条件:
  - system Python 自動検出は、bridge probe 成功済み候補だけを採用する
  - Debian 11 / Ubuntu 20.04 / WSL 想定の Python 3.8 / 3.9 + `python3-uno` false positive を regression test で再現し、以後は reject される
  - 明示 override も extraction 実行前に互換性エラーとして失敗し、`_run_bridge(...)` の `SyntaxError` まで遅延しない
  - smoke / docs / task 記述が「UNO import 可能」ではなく「bundled bridge 実行可能」で整合する
- 実装:
  - `_libreoffice_bridge.py` に internal `--probe` を追加し、`PropertyValue` と bridge 定数だけを解決する no-op 経路を用意した
  - `_python_supports_libreoffice_bridge(...)` は bundled bridge の `--probe` 実行で判定し、bundled / system Python 候補の両方で同じ互換性条件を使うようにした
  - `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` は同じ probe を必須にし、失敗時は `LibreOfficeUnavailableError` として fail-fast するようにした
  - `tests/conftest.py` の LibreOffice smoke gate も probe 失敗を unavailable 扱いに寄せた
  - README / 日本語 README / test requirements を probe ベースの説明へ更新した
- 追加検証:
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py tests/test_conftest_libreoffice_runtime.py -q` -> `22 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
