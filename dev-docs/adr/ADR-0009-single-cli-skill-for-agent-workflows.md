# ADR-0009: Single CLI Skill for Agent Workflows

## Status

`proposed`

## Background

ExStruct already exposes workbook editing through a stable local CLI:
`exstruct patch`, `exstruct make`, `exstruct ops`, and `exstruct validate`.
Those commands are intentionally positioned as the canonical local operational /
agent interface, while MCP remains the host-managed integration layer.

That still leaves a recurring operational problem for AI agents: commands alone
do not tell an agent when to use `patch` versus `make`, when to inspect op
schemas, when to run `validate` or `--dry-run`, how to respond to unsupported
ops, or when the request actually belongs to MCP because host-owned safety
controls are required.

The repository also needs one durable answer to two packaging questions that
would otherwise be re-decided repeatedly:

- whether ExStruct should ship one Skill or multiple narrower Skills for CLI
  editing
- whether the repository source of truth should live in the existing
  `.agents/skills/` tree or add a second in-repo mirror such as
  `.claude/skills/`

## Decision

- ExStruct adopts one installable Skill, `exstruct-cli`, for local AI-agent
  workflows around the existing editing CLI.
- The canonical repository source of truth for that Skill is
  `.agents/skills/exstruct-cli/`.
- The Skill consists of a lean `SKILL.md`, `agents/openai.yaml`, and focused
  reference documents under `references/`.
- `SKILL.md` carries only trigger-oriented rules, safety rules, the standard
  workflow, and direct navigation to the detailed references.
- Detailed operational knowledge such as command-selection guidance, risky-edit
  workflow, backend constraints, and verification guidance lives in
  `references/`.
- The Skill documents the CLI-versus-MCP boundary instead of collapsing both
  into one agent workflow:
  - local CLI Skill for local operational editing
  - MCP docs for host-owned path policy, transport, and artifact behavior
- This issue does not add `.claude/skills/` mirrors, sync automation, or
  Skill-local helper scripts/assets.

## Consequences

- ExStruct users and AI agents get one clear installation target for local CLI
  editing guidance instead of choosing among overlapping Skills.
- The repository gains a durable, reviewable place for agent workflow policy
  without changing the runtime CLI or MCP contracts.
- `SKILL.md` remains compact, which reduces trigger-time context cost and keeps
  detailed material available on demand through `references/`.
- The local CLI guidance and MCP host-policy guidance stay separated, which
  avoids teaching agents to bypass host-owned safety responsibilities.
- The chosen repo-only source-of-truth model keeps the repository simple, but it
  leaves packaging or mirroring into other agent-specific directories as a
  later concern.
- The initial implementation relies on lightweight validation and manual
  scenario review rather than a dedicated automated eval harness, so deeper
  agent-behavior testing remains future work if the Skill grows more complex.

## Rationale

- Tests:
  - `tests/cli/test_edit_cli.py`
  - `tests/cli/test_cli_lazy_imports.py`
  - Skill-specific automated tests do not exist yet; the initial verification
    path is `quick_validate.py` plus manual scenario review.
- Code:
  - `src/exstruct/cli/main.py`
  - `src/exstruct/cli/edit.py`
  - `.agents/skills/exstruct-cli/SKILL.md`
- Related specs:
  - `dev-docs/specs/editing-cli.md`
  - `dev-docs/specs/exstruct-cli-skill.md`
  - `docs/cli.md`
  - `docs/mcp.md`
  - `README.md`
  - `README.ja.md`

## Supersedes

- None

## Superseded by

- None
