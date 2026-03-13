---
name: adr-reconciler
description: Audit ExStruct ADRs against current specs, tests, and source code to detect policy drift, missing ADR updates, stale references, and evidence gaps. Use after merges, during periodic ADR audits, or when a review suspects that implementation and ADRs have diverged.
---

# ADR Reconciler

Audit the current decision records before proposing new text.

## Read

1. `dev-docs/agents/adr-governance.md`
2. `dev-docs/agents/adr-workflow.md`
3. `dev-docs/specs/adr-index.md`
4. Target ADRs under `dev-docs/adr/`
5. Related files under `dev-docs/specs/`, `tests/`, and `src/`

## Workflow

1. Select the target ADRs from changed domains, explicit ADR IDs, or related index entries.
2. Read each ADR claim and the linked specs, tests, and implementation paths.
3. Build an evidence matrix for every finding:
   - `adr`
   - `specs`
   - `src`
   - `tests`
4. Classify findings as one of:
   - `policy-drift`
   - `missing-adr-update`
   - `missing-evidence`
   - `stale-reference`
5. Recommend the next action:
   - `update-adr`
   - `new-adr`
   - `update-specs`
   - `add-tests`
   - `no-action`

## Output Contract

Return a concise audit result with:

- `scope`
  - `adrs`
  - `specs`
  - `src`
  - `tests`
- `findings`

Each finding should include:

- `type`
- `severity`
- `claim`
- `affected ADRs`
- `evidence matrix`
  - `adr`
  - `specs`
  - `src`
  - `tests`
- `recommended action`

Do not silently rewrite ADR text. If the finding implies a new or changed policy, route back through `adr-suggester` and `adr-drafter`.
