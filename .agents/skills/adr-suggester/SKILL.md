---
name: adr-suggester
description: Determine whether a change in ExStruct needs an ADR, using repository-specific governance and criteria. Use when reading an issue, PR, review thread, or diff and you need a verdict of required, recommended, or not-needed, plus candidate ADR titles and related existing ADRs.
---

# ADR Suggester

Classify the change before drafting anything.

## Read

1. `dev-docs/agents/adr-criteria.md`
2. `dev-docs/agents/adr-governance.md`
3. Relevant files under `dev-docs/adr/`

## Workflow

1. Read the issue, PR, diff, or review comment that introduced the change.
2. Read the relevant specs, implementation path, tests, and existing ADRs needed to justify the verdict.
3. Map the change to one or more ExStruct design domains.
4. Assemble an evidence triad:
   - specs contract
   - src symbol or execution path
   - tests that fix the current behavior, or an explicit note that evidence is missing
5. Return one verdict:
   - `required`
   - `recommended`
   - `not-needed`
6. If verdict is not `not-needed`, propose:
   - a candidate ADR title
   - existing ADRs to inspect or update
   - whether the next step is `new-adr` or `update-existing-adr`

## Output Contract

Return a concise result with:

- `verdict`
- `rationale`
- `affected domains`
- `existing ADR candidates`
- `suggested next action`
- `evidence triad`
  - `specs`
  - `src`
  - `tests`

Do not draft the ADR body.
