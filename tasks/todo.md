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
