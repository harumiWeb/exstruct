# ADR decision criteria

This document defines the criteria AI agents use to classify a change as `required`, `recommended`, or `not-needed`.

## Decision levels

### `required`

An ADR is required when any of the following apply:

- The contract or meaning of the public API, CLI, or MCP changes
- The meaning, requiredness, or omission policy of output JSON / YAML / TOON changes
- A default changes semantic contract, such as `default omit` vs `default include`
- Mode responsibility boundaries or validation contracts change
- The way the same constraints are aligned across API / CLI / engine / MCP entry points changes
- Backend addition, removal, priority, or fallback policy changes
- Runtime-failure handling, reason codes, log consistency, or return-shape combinations change
- A safety boundary such as `PathPolicy` changes
- Backward-compatibility policy changes
- Patch/edit backend capabilities or selection policy changes

### `recommended`

Changes that are likely to contain design decisions, but do not always require an ADR.

- Large revisions to the testing system
- Reorganization of error-handling policy
- Addition or replacement of a major dependency
- Changes to the basic strategy for performance optimization
- Turning operational flows or AI-agent procedures into permanent rules

### `not-needed`

ADR not needed because the change is not a policy-level change.

- typo fixes
- internal refactors with unchanged behavior
- wording cleanup that does not change the contract
- simple test additions that do not change design intent
- bug fixes that only restore the contract already defined by existing ADRs / specs / tests

## ExStruct-specific must-check areas

For the following areas, treat ADR-required as the first candidate as soon as a change appears.

| Area | Typical example | Existing references |
| --- | --- | --- |
| extraction mode | changing responsibilities of `light` / `libreoffice` / `standard` / `verbose` | `ADR-0001`, `dev-docs/specs/excel-extraction.md` |
| backend fallback | changing fallback contracts for COM / LibreOffice / future backends | `ADR-0002` |
| serialization contract | changing metadata omission/inclusion or schema meaning | `ADR-0003`, `docs/api.md`, `docs/cli.md` |
| patch backend policy | changing selection or constraints for `auto` / `com` / `openpyxl` | `ADR-0004`, `docs/mcp.md` |
| safety boundary | changing path allow ranges, normalization, or output-location rules | `ADR-0005`, `docs/mcp.md` |
| compatibility policy | changing preservation or break conditions for backward compatibility | existing ADR or a new ADR target |

## How to tell

If even one of the following questions is answered `yes`, strongly consider an ADR.

1. Is this a change in "why we adopt this policy" rather than "how we build it"?
2. Does the change affect multiple layers such as `docs/specs/tests/src`?
3. Is it likely that future contributors will face the same question again?
4. Does it break an assumption or exception condition from an existing ADR?
5. Is the likely review debate about which policy should be followed?

## Additional heuristics

- If a mode or backend change affects not just "where validation happens" but "whether the same validation is enforced at every entry point", it leans ADR-required
- For runtime-failure changes, consider reason codes, logs, and return-data shape as a three-part set
- A default change in serialization may look small, but it leans ADR-required because it changes the consumer contract
- For backend-selection changes, touching priority, forbidden fallback, or preflight capability checks makes the change lean ADR-required

## Cases where ADR is not needed but rationale should still be recorded

- regression fixes that only restore an existing ADR contract
- typo fixes in existing specs
- tests added according to an existing policy

In these cases, do not create a new ADR. Instead, leave a short rationale in the related issue, PR, `tasks/todo.md`, or review comment.

## Minimum output requirements for a verdict

At minimum, the AI agent must return:

- verdict: `required` / `recommended` / `not-needed`
- rationale: 1-3 lines of explanation
- affected domains: the related design areas
- existing ADR candidates: the existing ADRs to consult
- next action: `new-adr`, `update-existing-adr`, `no-adr`
- evidence triad:
  - contract text from specs
  - primary symbols or execution paths in src
  - fixed tests that lock the behavior down

## Phase 2: output requirements for audits and indexes

At minimum, `adr-reconciler` returns the following for drift audits of already-classified ADRs:

- scope:
  - target ADRs
  - reviewed `specs` / `src` / `tests`
- findings:
  - type: `policy-drift` / `missing-adr-update` / `missing-evidence` / `stale-reference`
  - severity: `high` / `medium` / `low`
  - claim
  - affected ADRs
  - evidence matrix:
    - ADR claim or relevant section
    - contract text from specs
    - primary symbols or execution paths in src
    - fixed tests that lock the behavior down
  - recommended action: `update-adr` / `new-adr` / `update-specs` / `add-tests` / `no-action`

At minimum, `adr-indexer` returns the following when updating indexes:

- updated artifacts
- added or changed ADR entries
- consistency findings

## Phase 3: output requirements for draft review

At minimum, `adr-reviewer` returns the following when reviewing ADR drafts:

- prerequisite: no unresolved `adr-linter` `high` / `medium` findings remain in the current draft
- verdict: `ready` / `revise` / `escalate`
- scope:
  - target ADR draft
  - related ADRs / `docs/` (when public API / CLI / MCP is involved) / `specs` / `src` / `tests`
  - referenced issue / PR / diff context
- findings:
  - type: `decision-gap` / `scope-conflict` / `evidence-risk` / `rollout-gap` / `ownership-escalation`
  - severity: `high` / `medium` / `low`
  - summary
  - why it matters
  - suggested revision
  - evidence:
    - relevant draft section or claim
    - related sources
- open questions
- residual risks
