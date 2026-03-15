# Feature Spec

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
