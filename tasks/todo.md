# Todo

## 2026-03-16 issue #99 phase 3 legacy monkeypatch compatibility follow-up

### Planning

- [x] `tasks/feature_spec.md` に legacy monkeypatch compatibility 修正方針を記録する
- [x] `patch_runner._sync_legacy_overrides()` で `patch_runner.get_com_availability` を `mcp.patch.runtime` まで同期する
- [x] `mcp.patch.service` を live module lookup ベースの legacy engine wrapper に切り替える
- [x] `tests/mcp/test_patch_runner.py` に `run_patch` / `run_make` の override 回帰を追加する
- [x] `tests/mcp/patch/test_service.py` に legacy engine monkeypatch 回帰を追加する
- [x] `tasks/lessons.md` に compat shim の monkeypatch 設計ルールを追記する
- [x] targeted pytest と `uv run task precommit-run` を実行する

### Review

- `src/exstruct/mcp/patch_runner.py` は `patch_runner.get_com_availability` を `mcp.patch.runtime` にも同期するようにし、`service._sync_compat_overrides()` を経由しても caller override が失われないようにした。
- `src/exstruct/mcp/patch/service.py` は `edit.engine.*` の copied alias ではなく、`mcp.patch.engine.*` を live module lookup する wrapper を `edit.service` に注入する形へ切り替えた。
- これにより `mcp.patch.service.apply_*` monkeypatch と `mcp.patch.engine.*.apply_*` monkeypatch の両方が `service.run_patch()` / `run_make()` 経路で有効になった。
- `tests/mcp/test_patch_runner.py` と `tests/mcp/test_make_runner.py` に entrypoint override precedence の回帰を追加し、`tests/mcp/patch/test_service.py` に legacy engine monkeypatch 回帰を追加した。
- `tasks/lessons.md` には compat shim で copied function alias を避け、public entrypoint で override precedence を回帰テスト化するルールを追加した。
- Verification:
  - `uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/patch/test_service.py -q`
  - `uv run pytest tests/mcp/patch/test_ops.py tests/mcp/patch/test_legacy_runner_ops.py -q`
  - `uv run pytest tests/edit -q`
  - `uv run task precommit-run`

## 2026-03-16 issue #99 phase 3 follow-up edit core decoupling from MCP implementation

### Planning

- [x] `tasks/feature_spec.md` に `src/exstruct/edit/**` から `exstruct.mcp.*` import を排除する follow-up 仕様を反映する
- [x] `edit.errors` / `edit.a1` / `edit.runtime` / `edit.engine.*` の ownership を edit 配下へ寄せ、edit から MCP import を除去する
- [x] `mcp.patch_runner` 側 path policy 経路を維持したまま、`edit.service` の `policy` 非依存化を完了する
- [x] `mcp.patch.internal` / `ops.*` / `runtime` / `service` を edit-backed compatibility path に整理する
- [x] core test を `tests/edit` に寄せ、MCP 側は shim / host behavior 回帰に絞る
- [x] `dev-docs/specs/editing-api.md`、`dev-docs/specs/data-model.md`、`dev-docs/architecture/overview.md`、`docs/mcp.md` を現行実装に合わせて更新する
- [x] `uv run pytest tests/edit -q` を実行する
- [x] `uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/patch -q` を実行する
- [x] `uv run task precommit-run` を実行する

### Review

- `src/exstruct/edit/**` から `exstruct.mcp.*` import は排除した。acceptance criteria の `rg -n "exstruct\\.mcp" src/exstruct/edit` は 0 件になった。
- `src/exstruct/edit/internal.py` を追加し、low-level patch implementation を edit-owned に移した。`edit.runtime` / `edit.service` / `edit.engine.*` はこの edit-owned implementation を使う。
- `src/exstruct/edit/service.py` は `policy` 非依存の pure core orchestration に戻し、MCP 側の path canonicalization は `src/exstruct/mcp/patch/service.py` の compatibility path で吸収した。
- `src/exstruct/mcp/patch/internal.py` は `exstruct.edit.internal` の typed compatibility shim に切り替え、repo 既存 tests が使う internal surface は維持した。
- `tests/edit/test_architecture.py` を追加し、`edit` package の import graph と fresh import での MCP side-effect 非依存を固定した。
- 恒久文書は `dev-docs/specs/editing-api.md`、`dev-docs/specs/data-model.md`、`dev-docs/architecture/overview.md`、`docs/mcp.md` に反映した。ADR verdict は継続して `not-needed`。
- Verification:
  - `uv run pytest tests/edit -q`
  - `uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/patch -q`
  - `uv run task precommit-run`

## 2026-03-16 pytest collect collision follow-up

### Planning

- [x] `tests/edit/test_service.py` と `tests/mcp/patch/test_service.py` の collect 衝突を再現確認する
- [x] edit 側 test module basename を一意な名前へ変更する
- [x] `tasks/lessons.md` に pytest test module naming の再発防止ルールを記録する
- [x] collect-only と関連 pytest を再実行して回帰がないことを確認する

### Review

- `uv run pytest tests/edit/test_service.py tests/mcp/patch/test_service.py --collect-only -q` で `import file mismatch` を再現し、指摘が妥当であることを確認した。
- `tests/edit/test_service.py` は `tests/edit/test_edit_service.py` に rename し、`tests/mcp/patch/test_service.py` との basename 衝突を解消した。
- 恒久ルールは `tasks/lessons.md` に移し、この section は session 記録として保持する。
- Verification:
  - `uv run pytest tests/edit/test_edit_service.py tests/mcp/patch/test_service.py --collect-only -q`
  - `uv run pytest tests/edit tests/mcp/patch -q`
  - `uv run task precommit-run`

## 2026-03-16 issue #99 phase 3 MCP rewiring to public edit core

### Planning

- [x] `exstruct.edit` を workbook editing の canonical core にする Phase 3 境界を固定する
- [x] `edit.models` を editing model の正本にし、`mcp.patch.models` を互換 shim に切り替える
- [x] `edit.runtime` に backend 選択・conflict handling・make seed orchestration を集約する
- [x] `edit.engine.*` を canonical engine boundary にし、`mcp.patch.engine.*` を互換 shim に寄せる
- [x] `edit.service` に patch/make orchestration を移し、`mcp.patch.service` を edit-backed wrapper に整理する
- [x] `mcp.patch_runner` で `PathPolicy` による path canonicalization と legacy monkeypatch override を吸収する
- [x] `mcp.tools` / `mcp.server` は tool payload、default `on_conflict`、artifact mirroring、thread offload など host 責務のみを維持する
- [x] core test と MCP shim test を分離して回帰を追加する
- [x] `dev-docs/specs/editing-api.md`、`dev-docs/specs/data-model.md`、`dev-docs/architecture/overview.md`、`docs/mcp.md` を現行実装に合わせて更新する
- [x] `uv run pytest tests/edit -q` を実行する
- [x] `uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py -q` を実行する
- [x] `uv run pytest tests/mcp/patch -q` を実行する
- [x] `uv run task precommit-run` を実行する

### Review

- `src/exstruct/edit/models.py` を canonical model 定義に切り替え、`src/exstruct/mcp/patch/models.py` は `edit.models` を再 export する shim に整理した。
- `src/exstruct/edit/runtime.py` と `src/exstruct/edit/engine/*` を canonical core とし、`src/exstruct/mcp/patch/runtime.py` と `src/exstruct/mcp/patch/engine/*` は互換 shim に寄せた。
- `src/exstruct/edit/service.py` に patch/make orchestration を置いたまま、public API の `policy` 非公開契約は `src/exstruct/edit/api.py` の wrapper で維持した。
- `src/exstruct/mcp/patch_runner.py` は `get_com_availability` monkeypatch を `edit.runtime` と legacy internal の両方へ同期する compatibility facade に整理した。
- `src/exstruct/mcp/patch/service.py` は legacy monkeypatch 先を `edit.service` / `edit.runtime` に伝播する wrapper として残し、既存 test monkeypatch 互換を維持した。
- `tests/edit/test_service.py` を追加し、COM 優先、auto fallback、make seed flow を core 観点で固定した。既存の `tests/mcp/*` は shim / host behavior 回帰として維持した。
- 恒久情報は `dev-docs/specs/editing-api.md`、`dev-docs/specs/data-model.md`、`dev-docs/architecture/overview.md`、`docs/mcp.md` に反映済みで、今回の task section は session 記録として保持する。
- ADR verdict は継続して `not-needed`。内部 ownership 移管であり、public contract / backend policy / safety boundary は変更していない。
- Verification:
  - `uv run pytest tests/edit -q`
  - `uv run pytest tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py -q`
  - `uv run pytest tests/mcp/patch -q`
  - `uv run task precommit-run`

## 2026-03-16 pr #103 unresolved review follow-up

### Planning

- [x] issue #99 の現行実装 PR を特定し、未解決 review thread を取得する
- [x] 指摘内容を現行コードと `editing-cli` 仕様に照らして妥当性確認する
- [x] `_load_patch_ops` の defensive guard を CLI 契約に沿う例外型へ揃える
- [x] 回帰テストを追加して `patch` / `make` の failure path を検証する
- [x] command-name collision で legacy extraction が壊れないよう dispatch を調整する
- [x] explicit edit syntax 判定が `--flag=value` form も扱えるようにする
- [x] targeted pytest と `uv run task precommit-run` を実行する

### Review

- `harumiWeb/exstruct` には PR `#99` は存在せず、issue `#99` の現行実装 PR `#103` を対象に確認した。着手時点の未解決 thread は 3 件で、いずれも妥当だった。
- `src/exstruct/cli/edit.py` の `_load_patch_ops()` だけ `TypeError` を投げると `patch` / `make` の clean error path から漏れる余地があったため、`ValueError` に統一した。
- `_load_patch_ops()` の defensive guard を `ValueError` に統一し、既存の `stderr` エラー + exit `1` 契約に載せた。
- `tests/cli/test_edit_cli.py` に helper 契約破壊を monkeypatch で注入する回帰テストを追加し、`patch` / `make` の両方で clean failure を確認した。
- `src/exstruct/cli/edit.py` の dispatch 判定は edit 固有シグナルがない限り既存ファイル名を優先するように調整し、`patch` / `make` / `ops` / `validate` と同名の legacy input を extraction path に戻した。
- `tests/cli/test_edit_cli.py` に command-name collision の回帰テストと、衝突時でも `--input` / `ops list` / `--help` などの explicit edit syntax が edit CLI を維持するテストを追加した。
- PR `#104` の追加 review 指摘に対応し、explicit edit syntax 判定は `--input=book.xlsx` / `--ops=ops.json` / `--input=...` の `--flag=value` form も edit CLI として扱うようにした。
- 永続化が必要な新規仕様や ADR 判断はなく、この task section は session 記録としてのみ保持する。
- Verification:
  - `uv run pytest tests/cli/test_edit_cli.py -q`
  - `uv run task precommit-run`

## 2026-03-15 tasks document cleanup

### Planning

- [x] `tasks/feature_spec.md` と `tasks/todo.md` の全 `##` section を棚卸しする
- [x] 既存 `dev-docs/`, `docs/`, `AGENTS.md`, ADR への吸収状況を確認する
- [x] `adr-suggester` と `adr-reconciler` の観点で ADR 追加要否を判定する
- [x] 過去 section を tasks から除去し、現行の cleanup 記録だけを残す
- [x] 差分と参照整合性を検証する

### Review

- `tasks/feature_spec.md` と `tasks/todo.md` の過去 section は全削除した。
- 恒久情報の正本は既存文書に揃っているため、ADR の新規追加・更新は不要だった。
- 主な恒久参照先: `dev-docs/adr/ADR-0001-extraction-mode-boundaries.md`, `dev-docs/adr/ADR-0002-rich-backend-fallback-policy.md`, `dev-docs/adr/ADR-0003-output-serialization-omission-policy.md`, `dev-docs/adr/ADR-0004-patch-backend-selection-policy.md`, `dev-docs/adr/ADR-0005-path-policy-safety-boundary.md`, `dev-docs/specs/excel-extraction.md`, `dev-docs/testing/test-requirements.md`, `dev-docs/agents/adr-governance.md`, `dev-docs/agents/adr-criteria.md`, `dev-docs/agents/adr-workflow.md`, `docs/api.md`, `docs/cli.md`, `docs/mcp.md`, `docs/release-notes/v0.6.1.md`, `CHANGELOG.md`, `AGENTS.md`
- `adr-suggester` verdict: `not-needed`
- `adr-reconciler` findings: なし
- Verification: `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
- Verification: `git diff --check -- tasks/feature_spec.md tasks/todo.md`

## 2026-03-15 issue #99 phase 1 public edit API

### Planning

- [x] issue #99 の Phase 1 境界を確認し、public API / MCP host boundary を整理する
- [x] `exstruct.edit` の公開面を定義する
- [x] op schema / normalize / type metadata の public import path を追加する
- [x] 既存 MCP compatibility path を維持する
- [x] ADR / internal spec / public docs を更新する
- [x] `uv run task precommit-run` を完走する

### Review

- `src/exstruct/edit/` を追加し、`patch_workbook` / `make_workbook` と patch 契約の公開 import path を導入した。
- `exstruct.edit` は `PathPolicy` を要求せず、既存の patch request/result 契約をそのまま利用する。
- `src/exstruct/mcp/patch/types.py`, `chart_types.py`, `specs.py`, `normalize.py`, `src/exstruct/mcp/op_schema.py` は `exstruct.edit` の契約モジュールを参照する互換 path に更新した。
- 恒久文書として `dev-docs/specs/editing-api.md` と `dev-docs/adr/ADR-0006-public-edit-api-and-host-boundary.md` を追加し、ADR index artifacts も同期した。
- Verification:
  - `uv run pytest tests/edit/test_api.py`
  - `uv run pytest tests/mcp/patch/test_normalize.py -q`
  - `uv run pytest tests/mcp/patch/test_service.py tests/mcp/test_patch_runner.py tests/mcp/test_make_runner.py -q`
  - `uv run task precommit-run`

## 2026-03-15 pr #102 review follow-up

### Planning

- [x] `normalize.py` の non-dict/non-str op handling を indexed `ValueError` に統一する
- [x] top-level `sheet` fallback が JSON-string op にも適用されるようにする
- [x] ADR-0006 の status を本文と index artifacts で同期する
- [x] review regression tests を追加する
- [x] PR 本文を template structure に合わせて更新する
- [x] targeted pytest と `uv run task precommit-run` を実行する

### Review

- `src/exstruct/edit/normalize.py` に raw op coercion helper を追加し、unsupported 型を indexed `ValueError` で拒否するようにした。
- `resolve_top_level_sheet_for_payload` は JSON-string op も dict 化してから alias 正規化と top-level `sheet` fallback を適用するようにした。
- ADR-0006 の status は `README.md`, `index.yaml`, `decision-map.md` を `accepted` に揃えた。
- review regression tests を `tests/mcp/patch/test_normalize.py`, `tests/mcp/test_tool_models.py`, `tests/mcp/test_server.py` に追加した。
- docstring warning 対応として、新規 `src/exstruct/edit/*.py` の不足 module docstring を補った。
- PR 本文は `.github/pull_request_template.md` の見出し構造に合わせて更新した。
- Verification:
  - `uv run pytest tests/mcp/patch/test_normalize.py tests/mcp/test_tool_models.py tests/mcp/test_server.py -q`
  - `uv run task precommit-run`

## 2026-03-15 pr #102 docs review follow-up

### Planning

- [x] unresolved review thread の内容を確認し、実文書との差分だけを直す
- [x] `data-model.md` の import path 表記を Python module path に揃える
- [x] `architecture/overview.md` の `edit/` tree を実ファイル構成に合わせる
- [x] 文書差分を確認し、必要なら PR thread を resolve する

### Review

- `dev-docs/specs/data-model.md` の “actual locations” は filesystem path ではなく Python import path を示す表現に修正した。
- `dev-docs/architecture/overview.md` の `edit/` tree に `chart_types.py` と `errors.py` を追加した。
- Verification:
  - `git diff --check`
  - `uv run task precommit-run`

## 2026-03-15 issue #99 phase 2 editing CLI

### Planning

- [x] issue #99 の Phase 2 を CLI 追加として固定し、legacy extraction CLI との互換境界を明文化する
- [x] `patch` / `make` / `ops list` / `ops describe` / `validate` の CLI surface を実装する
- [x] `patch` / `make` を `exstruct.edit` に接続し、JSON-first 出力と exit code 契約を整える
- [x] `validate` を入力ファイル検証 CLI として追加し、MCP `PathPolicy` なしで動作させる
- [x] CLI 回帰/新規テストを追加し、legacy extraction path の互換を確認する
- [x] ADR / internal spec / public docs / README.ja parity を更新する
- [x] targeted pytest と `uv run task precommit-run` を実行する

### Review

- `src/exstruct/cli/edit.py` を追加し、`patch` / `make` / `ops` / `validate` の editing CLI を実装した。
- `src/exstruct/cli/main.py` は first-token dispatch で editing subcommands だけを edit parser に回し、legacy extraction CLI はそのまま維持した。
- `patch` / `make` は `exstruct.edit` を呼び、`PatchResult` JSON を stdout に出したうえで `error is None` のときだけ exit `0` にした。
- `validate` は既存の入力ファイル検証ロジックを CLI に昇格し、`is_readable` / `warnings` / `errors` を JSON で返すようにした。
- 恒久文書として `dev-docs/specs/editing-cli.md` と `dev-docs/adr/ADR-0007-editing-cli-as-public-operational-interface.md` を追加し、ADR index artifacts と CLI/API/README 文書を更新した。
- Verification:
  - `uv run pytest tests/cli/test_edit_cli.py tests/cli/test_cli.py tests/cli/test_cli_alpha_col.py tests/edit/test_api.py tests/mcp/test_validate_input.py -q`
  - `uv run pytest tests/core/test_mode_output.py -q`
  - `uv run task precommit-run`
  - `git diff --check`
