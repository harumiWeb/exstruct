# ADR Review Contract

This document defines how `adr-reviewer` reviews ADR drafts and in what form it returns findings.

## Purpose

- Separate the structural inspection by `adr-linter` from the design review of ADR drafts
- Fix the review comment perspective so that "what counts as a problem" does not vary by AI
- Make explicit which issues the AI may fix and which must be escalated to a human

## Non-Goals

- Do not re-implement mechanical inspection of required sections or status as a substitute for `adr-linter`
- Do not perform post-merge drift auditing like `adr-reconciler`
- Do not make automated ADR body corrections the source of truth

## Preconditions

`adr-reviewer` is only valid when the current draft has been inspected by `adr-linter` and no unresolved `high` / `medium` findings remain.
When lint findings remain, fix the draft and re-run `adr-linter` before proceeding to design review.

## Review Focus

`adr-reviewer` checks the following perspectives in order.

1. decision coverage
   - Does the ADR actually resolve one policy decision?
   - Is `why` mixed with `how`?
2. scope and lineage
   - Are there conflicts, overlaps, or gaps with existing ADRs / specs?
   - Are issues that should use supersede or update being unnecessarily separated into a new ADR?
3. evidence strength
   - Do `Tests`, `Code`, and `Related specs` genuinely back up the decision?
   - Are consequences or claims extended without evidence?
4. rollout and compatibility
   - When public contracts, fallback, migration, or operational impact are relevant, are they addressed?
   - When the public API / CLI / MCP is involved, is the corresponding `docs/` in scope, and is there a basis for the compatibility / break judgment?
   - Are non-goals or reasons for not migrating omitted when they are needed?
5. ownership boundary
   - Are public API break judgments, security / license decisions, large-scale directory restructuring, or undecided product / spec policies being decided unilaterally by AI?

## Verdict

`adr-reviewer` verdicts are fixed to the following three values.

- `ready`
  - No high / medium severity design findings, and no unresolved issues outside AI's responsibility
- `revise`
  - Review findings remain that can be resolved by updating the ADR draft
- `escalate`
  - Issues remain that require human judgment and cannot be closed with the draft alone

`ready` is not a substitute for merge or final approval.

## Result Envelope

The review result must include at least the following top-level fields.

- `verdict`
- `scope`
- `findings`
- `open questions`
- `residual risks`

## Finding Contract

Return findings in severity order; each finding must include the following.

- `type`
- `severity`
- `summary`
- `why it matters`
- `suggested revision`
- `evidence`
  - `draft`
  - `related sources`

Finding types are fixed to the following.

- `decision-gap`
  - The ADR does not sufficiently resolve the policy question it should address
- `scope-conflict`
  - There is duplication, conflict, or mixed responsibilities with an existing ADR / spec
- `evidence-risk`
  - Cited evidence is weak, absent, or insufficient to support the claim
- `rollout-gap`
  - A necessary explanation of compatibility / migration / fallback / operational impact is missing
- `ownership-escalation`
  - The issue is outside AI's responsibility and requires human judgment

## Scope Contract

The review result must include at least the following in scope.

- Target ADR draft
- Investigated related ADRs
- Investigated related `docs/` (when the public API / CLI / MCP is involved)
- Investigated `dev-docs/specs/`, `src/`, `tests/`
- Referenced issue / PR / diff context

## Relationship to Other Skills

- Use after `adr-drafter`
- Use after `adr-linter` has cleared unresolved `high` / `medium` findings from the current draft
- When a public surface is involved, include the related `docs/` in the review scope
- If review results indicate policy-level drift, return to `adr-reconciler` or `adr-suggester`
