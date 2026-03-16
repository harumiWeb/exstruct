# Todo

## 2026-03-16 pr #103 unresolved review follow-up

### Planning

- [x] issue #99 の現行実装 PR を特定し、未解決 review thread を取得する
- [x] 指摘内容を現行コードと `editing-cli` 仕様に照らして妥当性確認する
- [x] `_load_patch_ops` の defensive guard を CLI 契約に沿う例外型へ揃える
- [x] 回帰テストを追加して `patch` / `make` の failure path を検証する
- [x] command-name collision で legacy extraction が壊れないよう dispatch を調整する
- [x] targeted pytest と `uv run task precommit-run` を実行する

### Review

- `harumiWeb/exstruct` には PR `#99` は存在せず、issue `#99` の現行実装 PR `#103` を対象に確認した。着手時点の未解決 thread は 3 件で、いずれも妥当だった。
- `src/exstruct/cli/edit.py` の `_load_patch_ops()` だけ `TypeError` を投げると `patch` / `make` の clean error path から漏れる余地があったため、`ValueError` に統一した。
- `_load_patch_ops()` の defensive guard を `ValueError` に統一し、既存の `stderr` エラー + exit `1` 契約に載せた。
- `tests/cli/test_edit_cli.py` に helper 契約破壊を monkeypatch で注入する回帰テストを追加し、`patch` / `make` の両方で clean failure を確認した。
- `src/exstruct/cli/edit.py` の dispatch 判定は edit 固有シグナルがない限り既存ファイル名を優先するように調整し、`patch` / `make` / `ops` / `validate` と同名の legacy input を extraction path に戻した。
- `tests/cli/test_edit_cli.py` に command-name collision の回帰テストと、衝突時でも `--input` / `ops list` / `--help` などの explicit edit syntax が edit CLI を維持するテストを追加した。
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
