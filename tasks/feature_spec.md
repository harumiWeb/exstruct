# Feature Spec

## 2026-03-18 pr #105 unresolved review follow-up

### Goal

- PR `#105` の現時点の未解決レビュー thread を再取得し、妥当なものだけを取り込む。
- `on_conflict="rename"` の atomic reservation で生じる zero-byte placeholder file が、openpyxl failure path で残らないようにする。
- legacy monkeypatch compatibility は既存 contract を維持し、妥当でない review は非採用理由を残す。

### Accepted findings

- `src/exstruct/edit/service.py` の `_apply_with_openpyxl()` は `ValueError` / `FileNotFoundError` / `OSError` を素通しで再送出しており、rename reservation 済み output path の zero-byte placeholder cleanup を行わない。
- `src/exstruct/mcp/patch/service.py` の `_sync_compat_overrides()` から `edit_service.apply_openpyxl_engine` / `apply_xlwings_engine` を差し替える経路は、Python の module global lookup により `edit.service` 内関数実行時にも有効であるため、「captured reference で monkeypatch が効かない」という指摘は現行コードでは不成立。

### Chosen constraints

- public API / CLI / MCP payload shape / backend selection policy は変更しない。
- cleanup 対象は rename reservation で作った zero-byte placeholder file のみとし、既存の user-managed file は削除しない。
- invalid 判定の review thread に合わせた実装修正は入れず、既存 monkeypatch regression test を再確認する。

### Test plan

- `tests/edit/test_edit_service.py` に、rename reservation 済み output path で `apply_openpyxl_engine()` が `ValueError` / `FileNotFoundError` / `OSError` を送出したとき placeholder file が cleanup される回帰を追加する。
- `tests/mcp/patch/test_service.py` の既存 legacy monkeypatch regression を再実行し、`mcp.patch.service` 指摘が再現しないことを確認する。

### ADR verdict

- `not-needed`
- rationale: cleanup bug fix と review validation であり、public contract や architecture policy の意味変更はない。

## 2026-03-17 pr #105 review follow-up

### Goal

- PR `#105` の未解決レビュー指摘を現行実装で再検証し、妥当なものだけを最小差分で取り込む。
- edit core の output-path / formula preflight / monkeypatch compatibility / runtime shim backward compatibility を維持する。
- public MCP guide は stable surface に寄せ、内部実装トポロジは `dev-docs/` に残す。

### Accepted findings

- `src/exstruct/edit/output_path.py` の `next_available_path()` は `exists()` 判定だけで rename 候補を返しており、競合時に同じ候補を返し得る。
- `src/exstruct/edit/output_path.py` の `next_available_directory()` は helper 単体では policy check より先に `_reserve_directory()` を実行し、副作用順序が逆転している。
- `src/exstruct/edit/service.py` の preflight formula error attribution は `typed_formula_issues[0]` を採っており、先頭が warning の場合に誤帰属し得る。
- `src/exstruct/mcp/patch_runner.py` は `edit.runtime` までは `get_com_availability` override を同期するが、`edit.internal` module-global は未同期で、`.xls` path の内部判定を取り逃す。
- `src/exstruct/mcp/patch/internal.py` と `src/exstruct/mcp/patch/runtime.py` は compatibility shim を名乗る一方で、legacy monkeypatch / `policy=` kwarg surface の一部を落としている。
- `src/exstruct/mcp/patch/service.py` の compatibility surface は実行上は動くが、legacy boundary より型情報が緩くなっている。
- `src/exstruct/edit/models.py` / `src/exstruct/edit/runtime.py` / `docs/mcp.md` / `tasks/todo.md` には ownership rename 後の stale wording が残っている。

### Chosen constraints

- public API / CLI / MCP payload shape / backend selection policy は変更しない。
- runtime shim の `policy` kwarg は legacy import path でのみ維持し、canonical `exstruct.edit.runtime` には持ち込まない。
- `next_available_path()` の原子的予約で生じる placeholder file は、dry-run や失敗時に cleanup して既存契約を崩さない。
- `tasks/todo.md` の旧ファイル名は、rename 再現手順として意味がある箇所は維持し、stale な成果記述だけ更新する。

### Test plan

- `tests/edit/test_edit_output_path.py` を追加し、rename path reservation と directory policy ordering を固定する。
- `tests/edit/test_edit_service.py` に preflight formula issue の warning/error 混在回帰を追加する。
- `tests/mcp/test_make_runner.py` に `patch_runner.get_com_availability` override が `.xls` make validation まで届く回帰を追加する。
- `tests/mcp/patch/test_legacy_runner_ops.py` に `mcp.patch.internal.get_com_availability` monkeypatch が legacy make path で効く回帰を追加する。
- `tests/mcp/patch/test_runtime_shim.py` を追加し、legacy runtime shim の `policy=` kwarg surface を固定する。

### ADR verdict

- `not-needed`
- rationale: compatibility bug fix と documentation cleanup であり、public contract や safety boundary の意味変更はない。

## 2026-03-16 issue #99 phase 3 legacy monkeypatch compatibility follow-up

### Goal

- `exstruct.edit` を canonical core のまま維持しつつ、`exstruct.mcp.patch_runner` / `exstruct.mcp.patch.service` の legacy monkeypatch 互換を修復する。
- `patch_runner.run_patch()` / `run_make()` が `patch_runner.get_com_availability` monkeypatch を実際の engine selection に反映するよう戻す。
- `mcp.patch.service.run_patch()` / `run_make()` が `mcp.patch.engine.*.apply_*` monkeypatch を経由して edit core に伝播するよう戻す。

### Accepted findings

- `patch_runner._sync_legacy_overrides()` は `mcp.patch.internal` の shim module と `edit.runtime` にしか `get_com_availability` を同期しておらず、直後の `mcp.patch.service._sync_compat_overrides()` が `patch.runtime.get_com_availability` で `edit.runtime` を上書きするため、`patch_runner.get_com_availability` monkeypatch が実効値として失われる。
- `mcp.patch.service._sync_compat_overrides()` は `exstruct.edit.engine.*` を `edit.service` に注入しているため、`mcp.patch.engine.openpyxl_engine.apply_openpyxl_engine` / `mcp.patch.engine.xlwings_engine.apply_xlwings_engine` を monkeypatch しても `service.run_patch()` / `run_make()` では反映されない。

### Chosen constraints

- `exstruct.edit` 配下へ `exstruct.mcp.*` import は追加しない。
- 互換優先順位は明示する。
  - `patch_runner.run_*` 経由では `patch_runner.get_com_availability` を canonical source とする。
  - `mcp.patch.service.run_*` を直接呼ぶ場合は `mcp.patch.runtime.get_com_availability` を canonical source とする。
- `mcp.patch.service.apply_*` monkeypatch 互換は維持したまま、`mcp.patch.engine.*.apply_*` monkeypatch も有効にする。
- public API / CLI / MCP payload shape / backend policy は変更しない。

### Test plan

- `tests/mcp/test_patch_runner.py` に、`patch_runner.run_patch()` / `run_make()` が `patch_runner.get_com_availability` monkeypatch を保持したまま edit core に到達する回帰を追加する。
- `tests/mcp/patch/test_service.py` に、`mcp.patch.engine.xlwings_engine.apply_xlwings_engine` と `mcp.patch.engine.openpyxl_engine.apply_openpyxl_engine` の monkeypatch が `service.run_patch()` で効く回帰を追加する。
- 既存の `service.apply_*` monkeypatch テストは維持し、両 compat surface の共存を確認する。

### ADR verdict

- `not-needed`
- rationale: 既存 compatibility shim の実装修正であり、public contract や safety boundary は変えない。

## 2026-03-16 issue #99 phase 3 follow-up edit core decoupling from MCP implementation

### Goal

- `src/exstruct/edit/**` から `exstruct.mcp.*` import を排除し、`exstruct.edit` を物理的にも workbook editing の canonical core にする。
- public API / CLI / MCP tool contract、backend selection/fallback policy、warning/error payload shape、`PathPolicy` safety boundaryは維持したまま、implementation ownership を `edit` 配下へ寄せる。
- MCP は host path policy と compatibility shim のみを持ち、editing behavior の source of truth にはならない。

### Current state

- `src/exstruct/edit/service.py` は canonical orchestration だが、`PathPolicy` 型を `exstruct.mcp.io` から import し、internal path canonicalization も受け持っている。
- `src/exstruct/edit/runtime.py` は `exstruct.mcp.patch.internal` を実効実装として参照し、`OnConflictPolicy` / `PathPolicy` も MCP 側から import している。
- `src/exstruct/edit/errors.py` は `PatchOpError` を `exstruct.mcp.patch.ops.common` から再 export している。
- `src/exstruct/edit/normalize.py` と `src/exstruct/edit/engine/*` はそれぞれ `exstruct.mcp.shared.a1` と `exstruct.mcp.patch.ops.*` を参照している。
- `src/exstruct/mcp/patch_runner.py` は compatibility facade だが、core 側がまだ MCP 実装に依存しているため、ownership が完全には反転していない。

### Chosen scope

- acceptance criteria は `rg -n "exstruct\\.mcp" src/exstruct/edit` が 0 件になることとする。
- `PatchOp`, `PatchRequest`, `MakeRequest`, `PatchResult` など public editing model の shape は変更しない。
- `edit.service.patch_workbook()` / `make_workbook()` は `policy` kwarg を廃止し、pure core orchestration に戻す。
- `PathPolicy` による root/deny_glob enforcement と request path canonicalization は `mcp.patch_runner` に集約する。
- `mcp.patch.internal`, `mcp.patch.ops.*`, `mcp.patch.runtime`, `mcp.patch.service`, `mcp.shared.*` は互換 import path を保つ shim として残してよいが、editing behavior の実体は持たない。

### Implementation decisions

- `src/exstruct/edit/errors.py` に `PatchOpError` の実体を置き、`src/exstruct/mcp/patch/ops/common.py` は re-export shim にする。
- `src/exstruct/edit/a1.py` に `parse_range_geometry` を含む edit 利用分の A1 helper を集約し、`edit.normalize` は edit-owned helper のみを使う。
- `src/exstruct/edit/runtime.py` は backend selection / fallback / conflict handling / make seed helper / supported extension / output path helper / large-op warning / range expansionを edit-owned 実装として持ち、`mcp.patch.internal` 依存を外す。
- `src/exstruct/edit/engine/openpyxl_engine.py` と `xlwings_engine.py` は edit-owned 実装を直接呼ぶ。必要なら `src/exstruct/edit/ops/` を新設し、既存 `mcp.patch.ops.*` から実装を移す。
- `src/exstruct/mcp/patch_runner.py` は `PathPolicy` を使って request path を許可済み絶対 path に正規化したうえで edit core を呼ぶ。`get_com_availability` monkeypatch も edit-owned runtime/helper に同期する。
- `src/exstruct/mcp/patch/service.py` / `runtime.py` / `engine/*` / `ops/*` / `internal.py` / `shared/a1.py` / `shared/output_path.py` は、repo 既存テストが依存している import path を維持する互換 shim へ寄せる。

### Contract invariants

- `exstruct.edit` の public import path と `patch_workbook(request)` / `make_workbook(request)` signature は維持する。
- CLI (`exstruct patch`, `exstruct make`, `exstruct ops`, `exstruct validate`) の引数・JSON 出力・exit code は変更しない。
- MCP tools (`exstruct_patch`, `exstruct_make`) の入出力 shape、tool 名、artifact mirroring、server default `on_conflict` は変更しない。
- backend selection / fallback policy と既存 warning / error payload shape は変更しない。
- repo 既存テストが使っている legacy monkeypatch points (`mcp.patch_runner.get_com_availability`, `mcp.patch.service.apply_*`) は維持する。

### Test and verification requirements

- `tests/edit` に、fresh import で `import exstruct.edit` が MCP side effect に依存しないことと、`src/exstruct/edit/**` に `exstruct.mcp` import が存在しないことを固定する static/architecture test を追加する。
- `tests/edit` は openpyxl/com/auto fallback/make seed flow/preflight/inverse ops と `PatchOpError` payload shape を core 観点で固定する。
- `tests/mcp/test_patch_runner.py` / `tests/mcp/test_make_runner.py` は `PathPolicy` root/deny_glob/relative path canonicalization と COM availability override sync を host behavior として維持する。
- `tests/mcp/patch/test_service.py` / `test_ops.py` / `test_models_internal_coverage.py` は legacy import path と monkeypatch compatibility が維持されることだけを確認する shim test に寄せる。
- 最終検証は `uv run pytest tests/edit -q`、`uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/patch -q`、`uv run task precommit-run` とする。

### ADR / docs retention

- ADR verdict は現時点で `not-needed`。ADR-0006 が既に `exstruct.edit` canonical core / MCP host boundary を記録しており、今回の変更はその implementation cleanup だからである。
- ただし public contract / backend policy / safety boundary / compatibility policy の意味変更が必要になった場合は ADR-0006 更新または新規 ADR を再判定する。
- 恒久文書の更新対象は `dev-docs/specs/editing-api.md`、`dev-docs/specs/data-model.md` Appendix A、`dev-docs/architecture/overview.md`、`docs/mcp.md` を最低ラインとする。

## 2026-03-16 issue #99 phase 3 MCP rewiring to public edit core

### Goal

- issue `#99` の Phase 3 として、MCP patch/make 実行経路を `exstruct.edit` 正本の editing core に再配線する。
- workbook editing の source of truth を `exstruct.edit` に寄せ、MCP は host/integration layer と互換 facade に責務を限定する。
- public API / CLI / MCP tool contract、`PathPolicy` safety boundary、backend selection/fallback policy は維持したまま内部 ownership を整理する。

### Current state

- `src/exstruct/edit/service.py` の `patch_workbook()` / `make_workbook()` は `exstruct.mcp.patch.service.run_patch()` / `run_make()` を `policy=None` で呼ぶだけの薄い wrapper である。
- `src/exstruct/mcp/tools.py` の `run_patch_tool()` / `run_make_tool()` は `PatchRequest` / `MakeRequest` を組み立てて `exstruct.mcp.patch_runner.run_patch()` / `run_make()` に委譲している。
- `src/exstruct/mcp/patch_runner.py` は `service.run_patch()` / `run_make()` に委譲する互換 facade であり、`get_com_availability` monkeypatch 伝播だけを自前で持っている。
- editing model の実体は依然として `src/exstruct/mcp/patch/models.py` にあり、`src/exstruct/edit/models.py` はそれを再 export している。
- `src/exstruct/mcp/patch/runtime.py` と `src/exstruct/mcp/patch/engine/*` は patch/make orchestration の実効依存として残っている。

### Chosen scope

- 今回の canonical core は `edit.models` / `edit.runtime` / `edit.engine.*` / `edit.service` までを対象とする。
- 低レベル op 実装 (`mcp.patch.internal`, `mcp.patch.ops.*`) は Phase 3 では全面移管しない。必要最小限の import 差し替えで edit-backed engine から再利用する。
- `mcp.patch.models` / `runtime` / `engine.*` / `service` / `patch_runner` は backward compatibility を維持する shim として残す。
- `mcp.tools` / `mcp.server` は host 責務のみを維持し、tool payload shape、artifact mirroring、thread offload、server default `on_conflict` を保持する。

### Contract invariants

- `exstruct.edit` の public import path は維持する。
- `PatchOp`, `PatchRequest`, `MakeRequest`, `PatchResult`, `PatchDiffItem`, `PatchErrorDetail`, `FormulaIssue` は `exstruct.edit` と `exstruct.mcp.patch_runner` の両方から引き続き import 可能にする。
- CLI (`exstruct patch`, `exstruct make`, `exstruct ops`, `exstruct validate`) の引数・JSON 出力・exit code は変更しない。
- MCP tools (`exstruct_patch`, `exstruct_make`) の入出力 shape、tool 名、`mirror_artifact`, `mirrored_out_path`, server default `on_conflict` 挙動は変更しない。
- `PathPolicy` は引き続き MCP host-owned behavior とし、`exstruct.edit` の public API には持ち込まない。
- backend selection / fallback policy と既存 warning / error payload shape は変更しない。

### Implementation decisions

- `src/exstruct/edit/models.py` に editing models の実体を移し、`src/exstruct/mcp/patch/models.py` は edit models を再 export する互換 module に変える。
- `src/exstruct/edit/runtime.py` を新設または同等の ownership 移管で、path-policy 非依存な runtime helper を保持する。MCP 固有の path 解決はここに置かない。
- `src/exstruct/edit/engine/openpyxl_engine.py` と `src/exstruct/edit/engine/xlwings_engine.py` を canonical engine boundary にし、既存 `mcp.patch.ops.*` を内部 backend adapter として呼ぶ。
- `src/exstruct/edit/service.py` が patch/make orchestration の正本となり、`src/exstruct/mcp/patch/service.py` は request を edit core に流す wrapper にする。
- `src/exstruct/mcp/patch_runner.py` は `PathPolicy` を使って request path を許可済み絶対 path に正規化し、その後で edit core を呼ぶ。`get_com_availability` の monkeypatch も edit runtime に同期する。
- `src/exstruct/mcp/tools.py` は request/result tool model の変換と artifact mirroring を担当し続けるが、editing behavior 自体は持たない。

### Test and verification requirements

- `tests/edit/test_api.py` は public API の成功経路に加えて、editing models の ownership が `edit` 側へ寄った後も import compatibility が保たれることを検証する。
- patch/make orchestration の主要回帰は `tests/edit` 側に移し、backend auto/com/openpyxl、fallback、formula preflight、make seed flow を core 観点で固定する。
- `tests/mcp/test_patch_runner.py` / `tests/mcp/test_make_runner.py` は `PathPolicy` による root/deny_glob と相対 path 解決、legacy monkeypatch override を固定する shim test として維持する。
- `tests/mcp/test_tools_handlers.py` / `tests/mcp/test_server.py` は tool payload 変換、default `on_conflict`、artifact mirroring、thread offload など host behavior の不変性を確認する。
- `tests/mcp/patch/test_models_internal_coverage.py` は `mcp.patch.models` が `edit.models` の互換 facade であることを確認する方向へ更新する。
- 最終検証は `uv run pytest tests/edit -q`、`uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py -q`、`uv run pytest tests/mcp/patch -q`、`uv run task precommit-run` とする。

### ADR / docs retention

- 現時点の ADR verdict は `not-needed`。理由は ADR-0006 が既に「public edit API を正本、MCP は host boundary」という方針を記録しており、Phase 3 はその内部実装寄せだからである。
- ただし implementation 中に public contract、backend policy、safety boundary、MCP payload shape のいずれかを変更する必要が出た場合は、ADR-0006 更新または新規 ADR を再判定する。
- 恒久文書の更新対象は `dev-docs/specs/editing-api.md`、`dev-docs/specs/data-model.md` Appendix A、`dev-docs/architecture/overview.md`、`docs/mcp.md` を最低ラインとする。

## 2026-03-16 pr #103 unresolved review follow-up

### Goal

- issue `#99` の現行 PR `#103` に残っている未解決 review thread を確認し、妥当なものだけを最小差分で取り込む。
- editing CLI の失敗契約を維持し、internal invariant break が起きても traceback ではなく `stderr` エラー + exit `1` に収束させる。
- legacy extraction entrypoint を維持し、bare token が edit subcommand 名と衝突する場合でも既存ファイル入力を優先できるようにする。

### Accepted finding

- `src/exstruct/cli/edit.py` の `_load_patch_ops()` は `resolve_top_level_sheet_for_payload()` の defensive guard で `TypeError` を投げており、`patch` / `make` の caller 側 catch 範囲から外れている。
- `dev-docs/specs/editing-cli.md` は JSON parse / validation / local I/O failure を CLI error + exit `1` と規定しており、同モジュール内の invariant guard だけ例外挙動が異なる理由はない。
- `src/exstruct/cli/main.py` の first-token dispatch は `patch` / `make` / `ops` / `validate` と同名の既存ファイル入力を edit CLI に誤送しうるため、legacy extraction compatibility の要件と衝突する。
- 既存回帰テストは `patch.xlsx` のような非衝突ケースしか見ておらず、command-name collision を防げない。
- explicit edit syntax の判定は `--input=book.xlsx` のような `--flag=value` 形式を見落としており、同名ファイルが存在すると valid な edit invocation を extraction に誤送しうる。

### Chosen constraints

- public CLI surface と exit-code policy は変更しない。
- defensive guard は `TypeError` 拡張ではなく `ValueError` に統一し、既存の user-facing error path に載せる。
- 回帰テストは monkeypatch で helper 契約破壊を注入し、`patch` と `make` の双方で clean failure を確認する。
- edit dispatch は bare token だけではなく edit 固有シグナルも見て判定し、既存ファイル名との衝突時は extraction を優先する。
- explicit edit syntax には `--flag value` だけでなく `--flag=value` の long-option form も含める。

## 2026-03-15 tasks document cleanup

### Goal

- `tasks/feature_spec.md` と `tasks/todo.md` から、完了済みの過去 task section を除去する。
- 恒久情報は `dev-docs/`、`docs/`、`AGENTS.md`、既存 ADR に残し、tasks には今回の棚卸し結果だけを残す。

### Retained references

- ADR: `dev-docs/adr/ADR-0001-extraction-mode-boundaries.md`, `dev-docs/adr/ADR-0002-rich-backend-fallback-policy.md`, `dev-docs/adr/ADR-0003-output-serialization-omission-policy.md`, `dev-docs/adr/ADR-0004-patch-backend-selection-policy.md`, `dev-docs/adr/ADR-0005-path-policy-safety-boundary.md`
- Internal specs: `dev-docs/specs/excel-extraction.md`, `dev-docs/specs/adr-index.md`, `dev-docs/specs/adr-review.md`, `dev-docs/testing/test-requirements.md`
- Agent governance: `dev-docs/agents/adr-governance.md`, `dev-docs/agents/adr-criteria.md`, `dev-docs/agents/adr-workflow.md`, `AGENTS.md`
- Public docs: `docs/api.md`, `docs/cli.md`, `docs/mcp.md`, `docs/release-notes/v0.6.1.md`, `CHANGELOG.md`

### ADR verdict

- `adr-suggester`: `not-needed`
- rationale: tasks 側に残っていたのは完了済みタスクの仕様メモと作業ログであり、既存恒久文書で説明できない policy-level decision は確認できなかった。
- `adr-reconciler` audit scope: `ADR-0001`〜`ADR-0005`, `dev-docs/specs/excel-extraction.md`, `dev-docs/testing/test-requirements.md`, `docs/api.md`, `docs/cli.md`, `docs/mcp.md`, `AGENTS.md`
- `adr-reconciler` findings: なし
- next action: `no-action`

## 2026-03-15 issue #99 phase 1 public edit API

### Goal

- Excel editing を `exstruct.edit` から利用できる first-class Python API として公開する。
- `PatchOp` / `PatchRequest` / `MakeRequest` / `PatchResult` と既存 op 契約を維持したまま、MCP 固有の path policy を public API から外す。
- MCP は互換レイヤとして維持し、tool I/O と host safety policy を担当し続ける。

### Public contract

- Primary public import path: `exstruct.edit`
- Public entry points:
  - `patch_workbook(request: PatchRequest) -> PatchResult`
  - `make_workbook(request: MakeRequest) -> PatchResult`
- Public helpers:
  - `coerce_patch_ops`
  - `resolve_top_level_sheet_for_payload`
  - `list_patch_op_schemas`
  - `get_patch_op_schema`
- Preserved Phase 1 contract:
  - existing op names remain unchanged
  - existing warning/error payload shapes remain unchanged
  - existing MCP compatibility imports remain valid

### Boundary

- `PathPolicy`, artifact mirroring, MCP tool payloads, server defaults, and thread offloading remain MCP-owned behavior.
- Phase 1 intentionally reuses the existing `exstruct.mcp.patch.*` execution pipeline under the hood to reduce backend regression risk while promoting the public API.

### Permanent references

- ADR: `dev-docs/adr/ADR-0006-public-edit-api-and-host-boundary.md`
- Internal specs:
  - `dev-docs/specs/editing-api.md`
  - `dev-docs/specs/data-model.md`
  - `dev-docs/architecture/overview.md`
- Public docs:
  - `docs/api.md`
  - `docs/mcp.md`

## 2026-03-15 pr #102 review follow-up

### Goal

- PR #102 の妥当な review 指摘だけを取り込み、Phase 1 の公開契約を変えずに不整合と正規化漏れを修正する。
- `coerce_patch_ops` / `resolve_top_level_sheet_for_payload` の JSON-string op 対応を、dict op と同じ indexed error shape で安定化する。
- ADR-0006 の status metadata を `accepted` に統一し、PR metadata warning は最小範囲で解消する。

### Accepted findings

- `coerce_patch_ops([None])` が `AttributeError` を漏らし、indexed validation error にならない。
- `resolve_top_level_sheet_for_payload` が JSON-string op を未解決のまま返し、top-level `sheet` fallback を適用できない。
- `ADR-0006` の本文と index artifacts の status が不一致。

### Chosen constraints

- public API signature と import path は変更しない。
- invalid op 型の失敗は `ValueError(build_patch_op_error_message(...))` に統一する。
- docstring warning 対応は、この PR で新規追加した `src/exstruct/edit/*.py` の不足 module docstring 補完までに限定する。
- PR 本文は `.github/pull_request_template.md` の見出し構造に合わせるが、Acceptance Criteria は issue #99 phase 1 用に書き換える。

## 2026-03-15 pr #102 docs review follow-up

### Goal

- PR #102 の docs review 指摘のうち、import path 表記の誤りと architecture tree の欠落だけを最小差分で修正する。
- 永続文書の説明を現行実装と一致させ、Phase 1 の API / runtime 契約自体は変更しない。

### Accepted findings

- `dev-docs/specs/data-model.md` の “actual locations” 先頭 bullet が import path と言いながら filesystem path を示している。
- `dev-docs/architecture/overview.md` の `edit/` tree に `chart_types.py` と `errors.py` が抜けている。

## 2026-03-15 issue #99 phase 2 editing CLI

### Goal

- Excel editing を first-class CLI として公開し、`exstruct.edit` を薄く包む operational interface を追加する。
- 既存の抽出 CLI `exstruct INPUT.xlsx ...` は互換維持し、編集系だけ subcommand を追加する。
- `PatchResult` と既存 patch op/schema 契約を崩さず、agent 向けに JSON-first な CLI を提供する。

### Public contract

- New CLI subcommands:
  - `exstruct patch`
  - `exstruct make`
  - `exstruct ops list`
  - `exstruct ops describe`
  - `exstruct validate`
- Chosen Phase 2 boundaries:
  - keep legacy extraction CLI entrypoint unchanged
  - do not add `exstruct extract` in this phase
  - `validate` means input-file readability validation, not patch request static validation
  - default output for new edit commands is JSON to stdout
- `patch` contract:
  - required flags: `--input`, `--ops`
  - optional flags: `--output`, `--sheet`, `--on-conflict`, `--backend`, `--auto-formula`, `--dry-run`, `--return-inverse-ops`, `--preflight-formula-check`, `--pretty`
  - `--ops` accepts a top-level JSON array from file or stdin (`-`)
  - exit `0` when `PatchResult.error is None`; otherwise emit serialized `PatchResult` and exit `1`
- `make` contract:
  - required flag: `--output`
  - optional flags: `--ops`, `--sheet`, `--on-conflict`, `--backend`, `--auto-formula`, `--dry-run`, `--return-inverse-ops`, `--preflight-formula-check`, `--pretty`
  - omitted `--ops` defaults to `[]`
- `ops` contract:
  - `list` returns compact JSON summaries (`op`, `description`)
  - `describe` returns detailed schema metadata for one op
- `validate` contract:
  - required flag: `--input`
  - output shape follows the existing input validation result (`is_readable`, `warnings`, `errors`)

### Implementation boundary

- `patch` / `make` / `ops` use `exstruct.edit` as the primary integration surface.
- `validate` may reuse the existing validation logic, but must not require MCP `PathPolicy`.
- Phase 2 does not change:
  - backend selection/fallback policy
  - patch result schema
  - MCP tool payloads or server safety policy
- Phase 2 also excludes:
  - backup / confirmation / allow-root / deny-glob flags
  - summary-mode output
  - request-envelope JSON input

### ADR verdict

- `adr-suggester`: `required`
- rationale: public CLI contract and CLI/API/MCP responsibility alignment change at policy level, while legacy extraction CLI compatibility is intentionally preserved.
- existing ADR candidates:
  - `ADR-0006-public-edit-api-and-host-boundary`
  - `ADR-0005-path-policy-safety-boundary`
  - `ADR-0004-patch-backend-selection-policy`
- suggested next action: `new-adr`
- candidate ADR title: `Editing CLI as Public Operational Interface`

### Permanent references

- ADR:
  - `dev-docs/adr/ADR-0007-editing-cli-as-public-operational-interface.md`
- Internal specs:
  - `dev-docs/specs/editing-api.md`
  - `dev-docs/specs/editing-cli.md`
  - `dev-docs/architecture/overview.md`
- Public docs:
  - `docs/cli.md`
  - `docs/api.md`
  - `README.md`
  - `README.ja.md`
