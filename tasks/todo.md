# Todo

## 2026-03-19 v0.7.0 release closeout

### Planning

- [x] Add the `0.7.0` changelog entry with `Added` / `Changed` / `Fixed`.
- [x] Create `docs/release-notes/v0.7.0.md` for issue `#99` and maintainer-facing docs changes.
- [x] Add `v0.7.0` to the `Release Notes` nav in `mkdocs.yml`.
- [x] Align the local package version in `pyproject.toml` and `uv.lock` to `0.7.0`.
- [x] Move the legacy monkeypatch compatibility precedence note to `dev-docs/architecture/overview.md`.
- [x] Compress `tasks/feature_spec.md` and `tasks/todo.md` to a single release-closeout section.
- [x] Run `uv run task build-docs`.
- [x] Run `uv run task precommit-run`.
- [x] Run the release-prep `rg` and `git diff --check` consistency checks.

### Review

- `CHANGELOG.md`, `docs/release-notes/v0.7.0.md`, `mkdocs.yml`, `pyproject.toml`, and `uv.lock` now describe and label the `0.7.0` release consistently.
- `dev-docs/architecture/overview.md` now holds the maintainer-facing note that compatibility shims must preserve live monkeypatch visibility and verify override precedence at the highest public entrypoint.
- Historical issue `#99`, PR follow-up, and prior cleanup logs were intentionally removed from `tasks/feature_spec.md` and `tasks/todo.md` after permanent information was classified and retained in `CHANGELOG.md`, `docs/`, and `dev-docs/architecture/`.
- No new `dev-docs/specs/` or `dev-docs/adr/` migration was required for this closeout; `ADR-0006`, `ADR-0007`, and the editing specs remain the canonical permanent sources.
- Verification:
  - `uv run task build-docs`
  - `uv run task precommit-run`
  - `rg -n "0\.7\.0|v0\.7\.0" pyproject.toml uv.lock CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.0.md`
  - `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
  - `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.0.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md dev-docs/architecture/overview.md`

## 2026-03-21 v0.7.1 release closeout

### Planning

- [x] Add the `0.7.1` changelog entry with `Added` / `Changed` / `Fixed`.
- [x] Create `docs/release-notes/v0.7.1.md` for issues `#107`, `#108`, and `#109`.
- [x] Add `v0.7.1` to the `Release Notes` nav in `mkdocs.yml`.
- [x] Align the local package version in `pyproject.toml` and the editable `exstruct` package entry in `uv.lock` to `0.7.1`.
- [x] Compress the detailed `#107` / `#108` working logs in `tasks/feature_spec.md` and `tasks/todo.md` into this release-closeout record.
- [x] Run `uv run pytest tests/cli/test_cli.py tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`.
- [x] Run `uv run task build-docs`.
- [x] Run `uv run task precommit-run`.
- [x] Run the release-prep `rg` and `git diff --check` consistency checks.

### Review

- `CHANGELOG.md`, `docs/release-notes/v0.7.1.md`, `mkdocs.yml`, `pyproject.toml`, and `uv.lock` now describe and label the `0.7.1` release consistently around the CLI/package import optimization work from issues `#107`, `#108`, and `#109`.
- The release narrative explicitly documents the public behavior deltas that shipped after `v0.7.0`: runtime validation for `--auto-page-breaks-dir`, lighter startup/import behavior for CLI and package entrypoints, preserved exported symbol names, and the restored `validate` error boundary.
- Historical implementation and review logs for issues `#107` and `#108` were intentionally removed from `tasks/feature_spec.md` and `tasks/todo.md` after permanent information was classified and retained in `CHANGELOG.md`, `docs/release-notes/v0.7.1.md`, `docs/cli.md`, `README.md`, `README.ja.md`, `dev-docs/specs/excel-extraction.md`, `dev-docs/architecture/overview.md`, and `ADR-0008`.
- No new `dev-docs/specs/` or `dev-docs/adr/` migration was required for this closeout; the existing CLI docs, architecture note, extraction spec, and `ADR-0008` remain the canonical permanent sources for the shipped behavior.
- Verification:
  - `uv run pytest tests/cli/test_cli.py tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`
  - `uv run task build-docs`
  - `uv run task precommit-run`
  - `rg -n "0\.7\.1|v0\.7\.1" CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.1.md`
  - `rg -n '^version = "0\.7\.1"$' pyproject.toml uv.lock`
  - `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
  - `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.1.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md`

## 2026-03-21 issue #115 ExStruct CLI Skill

### Planning

- [x] Confirm the ADR verdict and draft the new policy record for the CLI Skill packaging and CLI/MCP boundary decisions.
- [x] Add an internal spec under `dev-docs/specs/` for the `exstruct-cli` Skill contract.
- [x] Create `.agents/skills/exstruct-cli/` with `SKILL.md`, `agents/openai.yaml`, and the required `references/` files.
- [x] Keep `SKILL.md` concise and move detailed command, verification, and backend guidance into `references/`.
- [x] Update `README.md` and `README.ja.md` with Skill installation guidance, when-to-use guidance, and minimal example prompts.
- [x] Run the external `skill-creator` helper scripts to generate and validate `agents/openai.yaml`.
- [x] Run repo verification (`uv run task precommit-run`) and record the result.
- [x] Review which #115 notes remain temporary versus durable, then compress the `tasks/` sections after the durable content is migrated.

### Review

- Introduced the canonical repo-owned Skill at `.agents/skills/exstruct-cli/` with a lean `SKILL.md`, five focused reference documents, and generated `agents/openai.yaml`.
- Published `dev-docs/specs/exstruct-cli-skill.md` as the durable contract for the Skill layout, trigger/positioning rules, and verification obligations.
- Documented the policy decision in `dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md` and synchronized `dev-docs/adr/README.md`, `dev-docs/adr/index.yaml`, and `dev-docs/adr/decision-map.md`.
- Expanded `README.md` and `README.ja.md` with installation guidance, CLI-vs-MCP usage boundaries, and example prompts.
- The durable content for #115 now lives in `.agents/skills/exstruct-cli/`, `dev-docs/specs/exstruct-cli-skill.md`, `dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md`, `README.md`, and `README.ja.md`; this `tasks/` section remains only as a compact implementation record.
- Verification:
  - `python <skill-creator>/scripts/generate_openai_yaml.py .agents/skills/exstruct-cli --interface 'display_name=ExStruct CLI' --interface 'short_description=Guide safe ExStruct CLI edit workflows' --interface 'default_prompt=Use $exstruct-cli to choose the right ExStruct editing CLI command, follow a safe validate/dry-run workflow that inspects the PatchResult/diff before applying changes, and explain any backend constraints for this workbook task.'`
  - `python <skill-creator>/scripts/quick_validate.py .agents/skills/exstruct-cli`
  - `rg -n "ExStruct CLI Skill|exstruct-cli|validate -> dry-run -> inspect -> apply -> verify|\.agents/skills/exstruct-cli" README.md README.ja.md dev-docs/specs/exstruct-cli-skill.md dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md .agents/skills/exstruct-cli/ -g "*"`
  - `rg -n "^## |Tests:|Code:|Related specs:" dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md`
  - `uv run task precommit-run`
  - `git diff --check`
