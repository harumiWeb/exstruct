# Todo

## 2026-03-08 PR #76 unresolved thread follow-up

### Planning

- [x] PR #76 の未 resolve review thread を再取得して現行コードと突き合わせる
- [x] `tasks/feature_spec.md` に今回の follow-up contract を追記する
- [x] connector heuristic endpoint matching に rotation 反映を追加する
- [x] LibreOffice pipeline で rich artifact の部分成功を保持する
- [x] connector / pipeline fallback の regression tests を追加する
- [x] 関連 pytest と `uv run task precommit-run` で検証する

### Review

- PR #76 の未 resolve thread は 2026-03-08 時点で 2 件だった
  - `src/exstruct/core/backends/libreoffice_backend.py:539-541`
  - `src/exstruct/core/pipeline.py:1009-1010`
- 実装結果
  - `src/exstruct/core/backends/libreoffice_backend.py` は OOXML connector heuristic endpoint の `dx/dy` に `rotation` を反映してから begin/end 候補点を計算するようにした
  - `src/exstruct/core/pipeline.py` は LibreOffice rich backend 解決後の `extract_shapes()` と `extract_charts()` を分離し、chart failure 時は取得済み shape artifact を保持したまま fallback workbook を組み立てるようにした
  - `tests/core/test_libreoffice_backend.py` に回転付き heuristic endpoint regression test を追加した
  - `tests/core/test_pipeline_fallbacks.py` に `shapes success + charts failure` と `shapes failure short-circuit` の regression tests を追加した
- 検証結果
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py -q` -> `45 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`

## 2026-03-08 coverage / Codacy / CodeRabbit follow-up triage

### Planning

- [x] 最新 GitHub Actions run `22815422942` の failure reason を確認する
- [x] local non-COM coverage report を再実行し、coverage 低下の主戦場を特定する
- [x] PR #76 の最新 Codacy 指摘を再取得し、runtime defect と static-analysis false positive を切り分ける
- [x] PR #76 の最新 CodeRabbit / review thread を再確認し、未解決 thread と duplicate review comment を切り分ける
- [x] `tasks/feature_spec.md` に今回の accepted follow-up / out-of-scope を追記する
- [x] `engine.py` の per-call override 伝播を shared state mutation なしに組み直す
- [x] `engine.py` の immutability regression tests を追加する
- [x] LibreOffice startup 後の UNO bridge handshake を追加し、wrong-listener false positive を retry failure に戻す
- [x] wrong-listener / handshake failure の regression tests を追加する
- [x] `src/exstruct/core/libreoffice.py` の Codacy 対象 4 call site に最小スコープ suppression/comment を入れる
- [x] coverage 回復用の targeted tests を `libreoffice.py` / `libreoffice_backend.py` / `engine.py` に追加する
- [x] local non-COM suite coverage が `>= 80%` に戻ることを確認する

### Review

- 最新 GitHub Actions run は `2026-03-08` の `pytest` run `22815422942`
  - `test (ubuntu-latest, 3.12)` の `Run tests (non-COM suite)` が失敗
  - test 自体は `783 passed, 3 skipped, 11 deselected`
  - failure reason は `Coverage failure: total of 79 is less than fail-under=80`
  - Actions log の最終値は `Total coverage: 78.80%`
  - `ubuntu-latest, 3.11` / `windows-latest` matrix は fail-fast で `cancelled`
- ローカル再現:
  - `uv run pytest -m "not com and not render" --maxfail=1 --disable-warnings -q --cov=exstruct --cov-report=term-missing:skip-covered --cov-fail-under=0`
  - `785 passed, 1 skipped, 11 deselected`
  - coverage の低い変更対象は `src/exstruct/core/libreoffice.py (67%)`, `src/exstruct/core/backends/libreoffice_backend.py (79%)`, `src/exstruct/core/ooxml_drawing.py (83%)`
- Codacy (`python scripts/codacy_issues.py --pr 76 --min-level Warning`) は 7 件
  - `Bandit_B603` warning: `src/exstruct/core/libreoffice.py:267`, `:348`, `:398`
  - Semgrep error: `src/exstruct/core/libreoffice.py:267`, `:268`, `:348`, `:398`
  - 現行コードは `shell=False` + argv list + allowlisted env であり、今回は runtime exploit より analyzer 側の trust-boundary 誤検知として扱う
- CodeRabbit / review で現在も効いている追加指摘は 2 点
  - unresolved thread `src/exstruct/engine.py:250-265` (`also applies to 425-442`)
    - `_process_extract_scope()` が `self.output.destinations.auto_page_breaks_dir` を mutate し、immutable contract と caller isolation を壊す
  - latest review duplicate comment `src/exstruct/core/libreoffice.py:391-413` (`also applies to 768-835`)
    - `_reserve_tcp_port()` 後に別 listener が先に bind した場合でも `_wait_for_socket()` が startup success と誤認しうる
- 今回の方針
  - coverage gate は threshold 緩和ではなく targeted tests で戻す
  - `process()` の override 伝播は engine-level seam を維持しつつ explicit parameter 化し、shared state mutation を除去する
  - startup false positive は PID lookup ではなく UNO bridge handshake で詰める
  - Codacy は helper を前提に、該当 call site のみ inline suppression/comment で収束させる
- 実装結果
  - `src/exstruct/engine.py` は `_process_extract_scope()` を削除し、per-call destination がある場合だけ private override を `extract()` に渡す形へ変更した
  - `src/exstruct/core/libreoffice.py` は trusted subprocess helper と `--handshake` ベースの UNO startup 検証を追加し、wrong-listener を retry failure として扱う
  - `src/exstruct/core/_libreoffice_bridge.py` は `--handshake` / `--connect-timeout` を追加し、source file 自体の unit tests を追加した
- 検証結果
  - `uv run pytest tests/core/test_libreoffice_bridge.py tests/backends/test_auto_page_breaks.py tests/core/test_libreoffice_backend.py tests/engine/test_engine.py -q` -> `59 passed`
  - `uv run task precommit-run` -> `ruff`, `ruff-format`, `mypy` passed
  - `uv run pytest -m "not com and not render" --maxfail=1 --disable-warnings -q --cov=exstruct --cov-report=term-missing:skip-covered --cov-fail-under=80` -> `797 passed, 1 skipped, 11 deselected`, `Total coverage: 80.29%`

## 2026-03-08 PR #76 post-push triage

### Planning

- [x] 最新 GitHub Actions run / Codacy / review thread を再収集する
- [x] blocking issue と minor nit を切り分ける
- [x] `tasks/feature_spec.md` に post-push follow-up contract を追記する
- [x] `ExStructEngine.process(...)` が engine-level `extract(...)` seam を bypass しないように修正する
- [x] `tests/engine/test_engine.py::test_engine_process_normalizes_string_paths` の回帰を再現・固定する
- [x] `_read_relationships(...)` を typed relationship へ変更し、call site を relationship type 判定に切り替える
- [x] `tests/conftest.py` の LibreOffice runtime probe で broad `except Exception` をやめ、unexpected regression を surfacing する
- [x] `docs/api.md` の CLI 対応例に `--include-backend-metadata` を反映する
- [x] `_merge_anchor_geometry(...)` を anchor-first placement に揃え、位置マッチング回帰 test を追加する
- [x] `src/exstruct/core/libreoffice.py` の subprocess 呼び出しを trust-boundary 明示付き helper / suppression 方針で整理し、Codacy blocking issue を解消する

### Review

- 最新 GitHub Actions run は `2026-03-08` の `pytest` run `22814113410`
  - `test (ubuntu-latest, 3.12)` が `Run tests (non-COM suite)` で失敗
  - 実失敗は `tests/engine/test_engine.py::test_engine_process_normalizes_string_paths`
  - 例外は `ValueError: Excel file format cannot be determined, you must specify an engine manually.`
  - `windows-latest` jobs は fail-fast により `cancelled`
- Codacy は PR #76 に対して 4 件の `Error` を返している
  - 対象はすべて `src/exstruct/core/libreoffice.py`
  - ルールは `dangerous-subprocess-use-audit` と `dangerous-subprocess-use-tainted-env-args`
  - `shell=True` や command injection は現状見当たらないため、根本論点は subprocess の trust-boundary を code 上で明示できていない点と判断する
- 追加 review で妥当と判断した項目
  - `engine.process(...)` が engine-level `extract(...)` seam を bypass して test contract を壊している
  - `_read_relationships(...)` が `Type` を捨て、filename 推測に依存している
  - `tests/conftest.py` の broad `except Exception` が probe regression を skip に潰しうる
  - `docs/api.md` の CLI 例が `include_backend_metadata=True` と一致していない
  - `_merge_anchor_geometry(...)` が left/top を child-first のままにしている
- 今回 scope 外に置く項目
  - `pytest.MonkeyPatch` への import cleanup
  - `_start_soffice_startup_attempt(...)` の readability-only helper 分割
- 2026-03-08 実装・検証
  - `process()` は validation 後に engine-level `extract(...)` seam を通すよう修正し、string path 回帰 test を維持した
  - `_read_relationships(...)` は `Type` を保持する typed metadata へ変更し、sheet / drawing / chart 解決を relationship type 判定に切り替えた
  - `tests/conftest.py` は expected な unavailable failure だけを skip 扱いにし、unexpected probe regression は例外として surface させた
  - `docs/api.md` の CLI 相当例に `--include-backend-metadata` を反映した
  - `_merge_anchor_geometry(...)` は left/top を anchor-first に寄せ、回帰 test を追加した
  - LibreOffice subprocess は allowlist 環境変数と正規化済み executable/path helper を経由させ、bridge probe/env の回帰 test を追加した
  - `uv run pytest tests/engine/test_engine.py tests/test_conftest_libreoffice_runtime.py tests/core/test_libreoffice_backend.py -q` -> `44 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`

## 2026-03-08 PR #76 triage

### Planning

- [x] PR #76 の CI / Codacy / review 指摘を収集する
- [x] 各指摘を `feat/libreoffice-mode` の現行コードと GitHub ruleset に照らして妥当性判定する
- [x] follow-up に採用する項目を `tasks/feature_spec.md` に仕様として整理する
- [ ] `ExStructEngine.process(...)` と `extract(...)` の auto page-break override 不整合を修正する
- [ ] LibreOffice guardrail の test を拡張し、`image` / `.xls` / passthrough / extraction validator 分岐を直接カバーする
- [ ] `docs/cli.md` / `docs/api.md` の public surface 差分を解消する
- [ ] `confidence` の model constraint を追加し、schema 生成物と関連 test を同期する
- [ ] OOXML parsing hardening (`defusedxml`, anchor placement, scatter/bubble `xVal/yVal`) を実装する
- [ ] connector endpoint / direction の edge case を修正する
- [ ] LibreOffice session の `stderr` handling と smoke skip gate を harden する
- [ ] GitHub ruleset に `libreoffice-linux-smoke` を required status check として追加する
- [ ] OOXML-only shape/connector merge 方針と port reservation race の扱いを追加調査する

### Review

- 2026-03-08 時点の CI failure は `test (ubuntu-latest, 3.12)` の coverage 不足 (`78.62% < 80%`) であり、pytest 自体は `763 passed, 3 skipped, 11 deselected` だった
- Codacy は PR #76 に対して 1 件のみ `Error` を返し、対象は `src/exstruct/core/ooxml_drawing.py` の XML parser hardening だった
- 妥当と判断した主な指摘:
  - `process(..., auto_page_breaks_dir=...)` の override が extraction に伝播していない
  - guardrail test の分岐不足
  - `docs/cli.md` / `docs/api.md` の公開ドキュメントずれ
  - `libreoffice-linux-smoke` の skip 許容と ruleset 未設定
  - connector 片側未解決時の早期 return
  - `(dx, dy) = (0, 0)` の方向推定
  - `stderr=subprocess.PIPE` 未ドレイン
  - OOXML anchor placement 未使用
  - scatter/bubble `xVal` / `yVal` 未対応
  - `confidence` 範囲未制約
  - `xml.etree.ElementTree` の hardening 未対応
- 今回 scope 外とした主な指摘:
  - `ShapeData` / `ChartData` の dataclass 化
  - `load_workbook()` / `close_workbook()` typed handle 化
  - `normalize_path()` docstring のみの修正
  - `.xls` 例外型の `ConfigError` への変更

### 2026-03-08 Implementation Progress

- [x] `process()` / `extract()` の auto page-break override 整合を修正
- [x] LibreOffice guardrail tests (`image` / `.xls` / passthrough / extraction validator / process propagation) を追加
- [x] `docs/cli.md` / `docs/api.md` の follow-up docs 差分を反映
- [x] `confidence` の model/schema range 制約と validation/schema tests を追加
- [x] OOXML parser hardening (`defusedxml`) と anchor / scatter series fallback を実装
- [x] connector endpoint の mixed direct resolve と `(0,0)` direction fallback を修正
- [x] LibreOffice `stderr` を temp sink 化して undrained pipe を解消
- [x] LibreOffice smoke job を fail-fast 化し、GitHub ruleset に `libreoffice-linux-smoke` required status check を追加
- [ ] OOXML-only append 方針と `_reserve_tcp_port()` race は調査継続

### 2026-03-08 Verification Addendum

- `uv run pytest tests/test_constraints.py tests/backends/test_auto_page_breaks.py tests/core/test_libreoffice_backend.py tests/models/test_models_validation.py tests/models/test_schemas_generated.py tests/test_conftest_libreoffice_runtime.py tests/core/test_mode_output.py -k "libreoffice or process_excel_rejects or extract_workbook_rejects or not test_standard and not test_verbose" -q` -> `59 passed, 2 deselected`
- `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
- `gh api repos/harumiWeb/exstruct/rulesets/11087410` で `required_status_checks.context=libreoffice-linux-smoke` を確認

### 2026-03-08 Follow-up Planning

- [x] `_reserve_tcp_port()` race hardening を実装する
- [x] startup attempt 内の bounded retry contract を code / spec / tests に反映する
- [x] port collision を模した retry success / retry exhausted regression tests を追加する
- [x] draw-page snapshot あり sheet で unmatched OOXML shape / connector を append しない contract test を追加する
- [x] unmatched OOXML 件数の debug logging を追加する

### 2026-03-08 Follow-up Review

- `_reserve_tcp_port()` は hold-open reservation ではなく retry-based hardening で扱う
- OOXML-only append は v1 非対応とし、UNO canonical order を崩さない contract を固定する
- append を見送る代わりに、必要なら unmatched 件数の観測だけ別タスクで扱う
- `src/exstruct/core/backends/libreoffice_backend.py` で snapshot あり sheet の unmatched OOXML shape / connector 件数を `debug` logging するようにし、emit contract は変えずに観測だけ追加した
- 2026-03-08 follow-up implementation:
  - `src/exstruct/core/libreoffice.py` に startup strategy 内の bounded retry (`3` 回、short backoff) を追加し、失敗のたびに新しい port で `soffice` を起動し直すようにした
  - retry failure は `attempt N/3` 形式で集約し、strategy 間の failure aggregation と合わせて最終エラー文言に残すようにした
  - `tests/core/test_libreoffice_backend.py` に retry-within-strategy / retry-exhausted-to-shared-profile / aggregated-failure-detail の regression test を追加した
  - snapshot あり sheet では OOXML-only shape / connector を append しない contract test を追加し、v1 の emit 順序を固定した

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
## 2026-03-08 Linux LibreOffice CI smoke gate

### Planning

- [x] `tasks/feature_spec.md` に Linux CI smoke contract を追加し、`mode="libreoffice"` の required runtime gate を仕様化する
- [x] `.github/workflows/pytest.yml` に `libreoffice-linux-smoke` job を追加し、既存 unit matrix から分離する
- [x] `ubuntu-24.04` 上で `libreoffice` / `python3-uno` を導入して `RUN_LIBREOFFICE_SMOKE=1` で smoke test を実行する
- [x] smoke job は `tests/core/test_libreoffice_smoke.py` を対象にし、coverage upload とは分離する
- [x] README / README.ja / `docs/agents/TEST_REQUIREMENTS.md` に Linux required smoke job の運用を追記する
- [x] 変更内容の最終確認と検証コマンド結果を Review に反映する

### Review

- GitHub Actions に required Linux smoke job `libreoffice-linux-smoke` を追加し、`ubuntu-24.04` で LibreOffice runtime を導入して `pytest tests/core/test_libreoffice_smoke.py -m libreoffice` を実行する構成にした
- 既存の Linux/Windows unit matrix と Codecov upload はそのまま維持し、LibreOffice smoke は別 job に分離した
- `tasks/feature_spec.md` に Linux CI smoke contract を追加し、fallback 成功ではなく rich extraction 成功を merge 条件にする方針を明文化した
- README / 日本語 README / test requirements に required Linux smoke job と実行コマンドを追記した
- 追加検証:
  - `python -c "import yaml; yaml.safe_load(open('.github/workflows/pytest.yml', encoding='utf-8'))"` 相当の YAML parse -> `yaml-ok`
  - `uv run pytest tests/test_conftest_libreoffice_runtime.py -q` -> `1 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
