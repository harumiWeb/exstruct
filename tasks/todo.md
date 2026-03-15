# Todo

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
