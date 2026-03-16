# Feature Spec

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

### Chosen constraints

- public CLI surface と exit-code policy は変更しない。
- defensive guard は `TypeError` 拡張ではなく `ValueError` に統一し、既存の user-facing error path に載せる。
- 回帰テストは monkeypatch で helper 契約破壊を注入し、`patch` と `make` の双方で clean failure を確認する。
- edit dispatch は bare token だけではなく edit 固有シグナルも見て判定し、既存ファイル名との衝突時は extraction を優先する。

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
