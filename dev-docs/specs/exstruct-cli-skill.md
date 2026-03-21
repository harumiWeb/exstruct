# ExStruct CLI Skill Specification

This document defines the internal contract for the installable `exstruct-cli`
Skill that teaches AI agents how to use the existing ExStruct editing CLI.

## Purpose

- Package safe operational knowledge for local CLI-based workbook editing.
- Reuse the existing public editing CLI contract instead of creating a new
  runtime interface.
- Keep MCP host-policy concerns outside the Skill except where handoff guidance
  is needed.

## Non-goals

- Keep CLI, Python API, and MCP runtime behavior unchanged.
- Do not add `.claude/skills/` mirror or sync automation in the repository.
- Leave Skill-local `scripts/` and `assets/` optional in the initial
  implementation.

## Canonical repo layout

- Canonical source directory: `.agents/skills/exstruct-cli/`
- Required files:
  - `SKILL.md`
  - `agents/openai.yaml`
  - `references/command-selection.md`
  - `references/safe-editing.md`
  - `references/ops-guidance.md`
  - `references/verify-workflows.md`
  - `references/backend-constraints.md`

## Trigger and positioning contract

- The Skill name is `exstruct-cli`.
- The frontmatter `description` must make the Skill trigger for:
  - Excel workbook create/edit requests through ExStruct CLI
  - `patch` versus `make` selection
  - `validate` / `dry-run` / `ops describe` guidance
  - safe CLI workflows for AI agents
- The Skill must position interfaces consistently with the public docs:
  - use the editing CLI for local operational / agent workflows
  - use MCP when host-owned path policy, transport, or artifact behavior is
    required
  - use `openpyxl` / `xlwings` for ordinary imperative Python editing

## `SKILL.md` contract

- Keep `SKILL.md` concise and workflow-oriented.
- Include only:
  - YAML frontmatter with `name` and `description`
  - command-selection rules
  - safety rules
  - standard workflow steps
  - minimal examples
  - direct links to each `references/` file
- Do not place large op catalogs, backend deep-dives, or long recipe
  collections in `SKILL.md`.

## `references/` contract

### `references/command-selection.md`

- Map user intents to `exstruct patch`, `exstruct make`, `exstruct validate`,
  `exstruct ops list`, and `exstruct ops describe`.
- Define when the request should be redirected to MCP guidance instead of the
  local CLI Skill.

### `references/safe-editing.md`

- Define the standard flow:
  `validate -> dry-run -> inspect -> apply -> verify`
- Explain ambiguous-request handling and destructive-edit safeguards.
- Explain why `backend=openpyxl` is the safe default for same-engine dry-run
  comparisons.

### `references/ops-guidance.md`

- Explain how to inspect supported ops before editing.
- State that unsupported ops must not be invented.
- Describe how to respond when the requested capability is missing or unclear.

### `references/verify-workflows.md`

- Specify what to inspect in `PatchResult`.
- Establish when re-validation or re-extraction is required.
- Clarify when lightweight verification is acceptable.

### `references/backend-constraints.md`

- Summarize `openpyxl`, `com`, and `auto` behavior relevant to agent guidance.
- Cover `.xls` and `create_chart` constraints.
- State explicit failure behavior for unsupported requests.

## `agents/openai.yaml` contract

- Must be present for repo consistency with other checked-in Skills.
- Must contain:
  - `interface.display_name`
  - `interface.short_description`
  - `interface.default_prompt`
- Generate or refresh it with the external `skill-creator` helper script rather
  than hand-maintaining inconsistent values.

## Public docs obligations

- `README.md` and `README.ja.md` must both:
  - describe the Skill as one installable entry point
  - explain the preferred one-command install path via
    `npx skills add harumiWeb/exstruct/.agents/skills --skill exstruct-cli`
  - explain the manual fallback for unpublished branches or runtimes that do
    not support `npx skills add`
  - describe when to use the Skill versus MCP documentation
  - provide at least one minimal example prompt

## Verification obligations

- Run verification from the repository root with explicit commands against
  `.agents/skills/exstruct-cli`:
  - `python <skill-creator>/scripts/generate_openai_yaml.py .agents/skills/exstruct-cli --interface ...`
  - `python <skill-creator>/scripts/quick_validate.py .agents/skills/exstruct-cli`
- Run `uv run task precommit-run`.
- Manually review representative scenarios for:
  - create versus edit routing
  - unknown-op handling
  - risky-edit safety flow
  - backend-constraint guidance
  - CLI-versus-MCP boundary guidance
