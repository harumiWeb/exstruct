---
name: adr-reviewer
description: Review an ExStruct ADR draft for decision quality, overlap with existing ADRs and specs, evidence strength, rollout risk, and human-ownership escalations. Use only after adr-linter reports no unresolved high/medium findings on the current draft, and when you need design-review findings before merge or handoff.
---

# ADR Reviewer

Review the policy decision, not just the document shape.

## Read

1. `dev-docs/agents/adr-governance.md`
2. `dev-docs/agents/adr-criteria.md`
3. `dev-docs/agents/adr-workflow.md`
4. `dev-docs/specs/adr-review.md`
5. The target ADR draft
6. Related ADRs, relevant public docs under `docs/` when public API / CLI / MCP contracts are in scope, internal specs, tests, src paths, and issue / PR context
7. Existing `adr-linter` findings for the current draft

## Workflow

1. Confirm the current draft has no unresolved `adr-linter` `high` / `medium` findings. Only proceed with design review after that precondition is met.
2. Read the ADR draft and identify the single policy question it is trying to resolve.
3. Check whether the draft overlaps with, contradicts, or should supersede an existing ADR or spec.
4. Verify that the cited `Tests`, `Code`, and `Related specs` actually support the claims being made, and include relevant public `docs/` pages in scope when the ADR touches public API / CLI / MCP contracts.
5. Review whether compatibility, rollout, fallback, migration, or safety consequences are covered when relevant.
6. Detect human-owned decisions that AI should not settle, including public API break judgment, security or license calls, major directory reorganization, or unresolved product/spec direction.
7. Return one verdict:
   - `ready`
   - `revise`
   - `escalate`

## Output Contract

Return findings first, ordered by severity, and include:

- `verdict`
- `scope`
  - `draft`
  - `related ADRs`
  - `public docs`
  - `specs`
  - `src`
  - `tests`
  - `issue / PR context`
- `findings`

Each finding should include:

- `type`
  - `decision-gap`
  - `scope-conflict`
  - `evidence-risk`
  - `rollout-gap`
  - `ownership-escalation`
- `severity`
- `summary`
- `why it matters`
- `suggested revision`
- `evidence`
  - `draft`
  - `related sources`

Also include top-level:

- `open questions`
- `residual risks`

Do not silently rewrite the ADR text. If the review hits a human-owned decision, return `escalate` instead of inventing a final policy.
