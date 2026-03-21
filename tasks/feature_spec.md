# Feature Spec

## 2026-03-19 v0.7.0 release closeout

### Goal

- Publish the `v0.7.0` release-prep artifacts for the workbook editing work delivered through issue `#99`.
- Collapse temporary issue and review logs after moving the only remaining maintainer-facing rule into permanent documentation.
- Keep a compact closeout record that states where permanent information now lives and how the release-prep work was verified.

### Permanent destinations

- `CHANGELOG.md`
  - Holds the `0.7.0` `Added` / `Changed` / `Fixed` summary for the public release.
- `docs/`
  - `docs/release-notes/v0.7.0.md` records the user-facing release narrative for issue `#99`, review follow-ups, and maintainer-facing documentation work.
  - `mkdocs.yml` keeps the canonical `Release Notes` navigation entry for `v0.7.0`.
- `dev-docs/architecture/overview.md`
  - Records the legacy monkeypatch compatibility precedence note for compatibility shims.
- `dev-docs/specs/`
  - No new spec migration was required for this closeout; the existing editing API and CLI specs remain the canonical contract source.
- `dev-docs/adr/`
  - No new ADR or ADR update was required; existing `ADR-0006` and `ADR-0007` continue to cover the policy boundary.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain only this release-closeout record plus verification, not the historical phase-by-phase work log.

### Constraints

- Public API / CLI / MCP contracts and backend selection policy remain unchanged in this closeout.
- `README.md` and `docs/index.md` do not gain direct release-note links; `mkdocs.yml` stays the canonical route.
- `uv.lock` is not fully regenerated; only the local `exstruct` package version is aligned to `0.7.0`.

### Verification

- `uv run task build-docs`
- `uv run task precommit-run`
- `rg -n "0\.7\.0|v0\.7\.0" pyproject.toml uv.lock CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.0.md`
- `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
- `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.0.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md dev-docs/architecture/overview.md`

### ADR verdict

- `not-needed`
- rationale: this was release preparation and task-log retention cleanup. The policy decisions already live in `ADR-0006`, `ADR-0007`, and the editing specs.

## 2026-03-21 v0.7.1 release closeout

### Goal

- Publish the `v0.7.1` release-prep artifacts for the CLI and package import startup optimization work delivered through issues `#107`, `#108`, and `#109`.
- Collapse the temporary issue and review logs for `#107` and `#108` after confirming that the durable contract and design rationale already live in permanent documentation.
- Keep a compact closeout record that states where permanent information now lives and how the release-prep work was verified.

### Public contract summary

- `--auto-page-breaks-dir` is always listed in extraction CLI help output and validated only when the flag is requested at runtime.
- `exstruct --help`, `exstruct ops list`, non-edit CLI routing, `import exstruct`, and `import exstruct.engine` now defer heavy imports until execution actually needs them.
- Public exported symbol names from `exstruct` and `exstruct.edit` remain stable; only import timing changed.
- The edit CLI `validate` subcommand keeps its narrow historical error boundary and must still propagate `RuntimeError`.
- No new CLI commands, MCP payload shapes, or backend-selection policy changes are introduced in this closeout.

### Permanent destinations

- `CHANGELOG.md`
  - Holds the `0.7.1` `Added` / `Changed` / `Fixed` summary for the public release.
- `docs/`
  - `docs/release-notes/v0.7.1.md` records the user-facing release narrative for issues `#107`, `#108`, and `#109`.
  - `mkdocs.yml` keeps the canonical `Release Notes` navigation entry for `v0.7.1`.
  - `docs/cli.md` remains the canonical public CLI contract for extraction help/runtime validation behavior.
- `README.md` and `README.ja.md`
  - Retain the public-facing wording for extraction runtime validation and CLI behavior that shipped with issue `#107`.
- `dev-docs/specs/`
  - `dev-docs/specs/excel-extraction.md` remains the canonical internal guarantee for extraction CLI runtime validation.
- `dev-docs/architecture/overview.md`
  - Records the durable lightweight-startup rule for package `__init__` files, CLI routing, and `exstruct.engine`.
- `dev-docs/adr/`
  - `ADR-0008` remains the canonical policy source for runtime capability validation in the extraction CLI.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain only the release-closeout records plus verification, not the detailed issue-by-issue implementation log.

### Constraints

- `README.md` and `docs/index.md` do not gain direct release-note links; `mkdocs.yml` stays the canonical navigation route.
- `uv.lock` is not fully regenerated; only the editable `exstruct` package version is aligned to `0.7.1`.
- This closeout does not add a new ADR or new permanent spec document; it only points to the existing permanent sources for the shipped behavior.

### Verification

- `uv run pytest tests/cli/test_cli.py tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`
- `uv run task build-docs`
- `uv run task precommit-run`
- `rg -n "0\.7\.1|v0\.7\.1" CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.1.md`
- `rg -n '^version = "0\.7\.1"$' pyproject.toml uv.lock`
- `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
- `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.1.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md`

### ADR verdict

- `not-needed`
- rationale: this was release preparation and task-log retention cleanup. The shipped policy decisions already live in `ADR-0008`, the extraction docs/specs, and the architecture note.

## 2026-03-21 issue #115 ExStruct CLI Skill

### Goal

- Add one installable Skill that teaches AI agents how to use the existing ExStruct editing CLI safely and consistently.
- Keep the current CLI, Python API, and MCP runtime contracts unchanged while packaging operational knowledge into repo-owned skill assets and durable documentation.
- Record the long-lived policy choices behind the Skill structure, CLI/MCP boundary, and distribution model in permanent documents before implementation logs are compressed.

### Public contract summary

- No changes to `exstruct patch`, `exstruct make`, `exstruct ops`, `exstruct validate`, `exstruct.edit`, or MCP tool payloads.
- Public docs gain one new documented interface layer: a single installable Skill named `exstruct-cli` for agent-side CLI workflows.
- `README.md` and `README.ja.md` must describe:
  - where the Skill source lives in the repository
  - how to install/copy it into an agent runtime
  - when to use the Skill versus MCP guidance
  - minimal usage examples / prompts

### Internal Skill contract

- Canonical repo source: `.agents/skills/exstruct-cli/`
- Required files:
  - `SKILL.md`
  - `agents/openai.yaml`
  - `references/command-selection.md`
  - `references/safe-editing.md`
  - `references/ops-guidance.md`
  - `references/verify-workflows.md`
  - `references/backend-constraints.md`
- `SKILL.md` must stay lean and contain only trigger-oriented frontmatter, command selection rules, safety rules, workflow steps, and direct links to `references/`.
- `references/` must hold the detailed operational knowledge and avoid duplicating long catalogs inside `SKILL.md`.
- `scripts/` and `assets/` are out of scope for the initial implementation unless a concrete deterministic need appears during writing or validation.

### Permanent destinations

- `.agents/skills/exstruct-cli/`
  - Canonical source for the installable Skill assets.
- `README.md` and `README.ja.md`
  - Public installation and usage guidance for the Skill.
- `dev-docs/specs/`
  - `dev-docs/specs/exstruct-cli-skill.md` is the canonical internal spec for the ExStruct CLI Skill contract and required reference structure.
- `dev-docs/adr/`
  - `ADR-0009-single-cli-skill-for-agent-workflows.md` records the policy-level decisions: single Skill, repo-owned source of truth, and CLI-versus-MCP documentation boundary.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Keep the working record only until the durable information above is written.

### Constraints

- Use the existing repository convention for skills (`.agents/skills/` + `agents/openai.yaml`); do not add `.claude/skills/` mirrors or sync automation in this issue.
- Keep the Skill focused on local CLI workflows; MCP host-policy and transport concerns stay documented separately under the existing MCP docs.
- Follow docs parity rules: any public README change in English must be mirrored in `README.ja.md` in the same change.
- Reuse the external `skill-creator` helper scripts for `openai.yaml` generation and lightweight validation rather than adding new repo-local validation code for this issue.

### Verification

- `python <skill-creator>/scripts/generate_openai_yaml.py .agents/skills/exstruct-cli --interface ...`
- `python <skill-creator>/scripts/quick_validate.py .agents/skills/exstruct-cli`
- `uv run task precommit-run`
- Manual scenario review of:
  - create-vs-edit command selection
  - `validate -> dry-run -> inspect -> apply -> verify` guidance
  - unsupported-op / backend-constraint handling
  - CLI-vs-MCP routing guidance
  - README English/Japanese parity

### ADR verdict

- `recommended`
- rationale: the change turns AI-agent operational workflow into a durable repository rule and resolves recurring tradeoffs around single-skill packaging, repo source of truth, and the CLI-versus-MCP boundary.
