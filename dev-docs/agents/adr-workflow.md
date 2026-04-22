# ADR workflow

This document defines the standard flow for handling ADRs from issues and PRs.

## Scope of Phase 1

Phase 1 standardizes only the following:

1. Determining whether an ADR is needed
2. Drafting a new ADR or proposing an update to an existing ADR
3. Linting ADR documents

## Additions in Phase 2

Phase 2 adds the following on top of Phase 1:

1. Auditing ADR consistency against `specs` / `tests` / `src` via `adr-reconciler`
2. Updating ADR indexes and relationship maps via `adr-indexer`

## Additions in Phase 3

Phase 3 adds the following on top of Phase 1 + 2:

1. Design review of ADR drafts via `adr-reviewer`

## Standard flow

1. Read the issue or PR
2. Read related `docs/`, `dev-docs/specs/`, `dev-docs/adr/`, `tests/`, and `src/`, and gather the evidence triad needed for the decision
3. Use `adr-suggester` to decide `required` / `recommended` / `not-needed`
4. Even when the verdict is `not-needed`, leave the rationale and evidence triad in the issue or PR
5. When the verdict is `required` or `recommended`, use `adr-drafter` to produce a new ADR draft or a proposal to update an existing ADR
6. Use `adr-linter` to check structure and evidence
7. If `adr-linter` reports unresolved `high` / `medium` findings, revise the draft and rerun step 6
8. Use `adr-reviewer` to review design soundness, conflicts with existing ADRs, and compatibility / rollout / fallback / safety impact. If the ADR touches public API / CLI / MCP, include related `docs/` in scope
9. If `adr-reviewer` returns `revise`, revise the draft and rerun steps 6-8 as needed
10. If `adr-reviewer` returns `escalate`, send the issue back to the issue or PR as a point that requires human judgment
11. Run `adr-reconciler` when an ADR is newly added or updated, or when the change includes a policy-level shift
12. At merge time, recheck consistency with related specs / docs / tests and reconciliation findings
13. If an ADR was added, updated, or superseded, run `adr-indexer` to synchronize `README.md`, `index.yaml`, and `decision-map.md`

## Reading order

For ADR-related tasks, review materials in this order:

1. `docs/`
2. `dev-docs/specs/`
3. `dev-docs/adr/`
4. `tests/`
5. `src/`

Only when AI-oriented decision guidance is needed, also read:

- `dev-docs/agents/adr-governance.md`
- `dev-docs/agents/adr-criteria.md`

## Responsibility of each skill

### `adr-suggester`

- Decide whether a change should be treated as a design decision
- Gather the evidence triad before returning a verdict
- Return new-ADR candidates and existing-ADR candidates
- Include the evidence triad even for `not-needed`
- Do not generate ADR body text

### `adr-drafter`

- Create either a new ADR draft or a proposal to update an existing ADR
- Fill in `Background`, `Decision`, `Consequences`, and `Rationale`
- Include `Tests`, `Code`, and `Related specs` in the `Rationale` section

### `adr-linter`

- Check `Status`, required sections, evidence, and `Supersedes` / `Superseded by`
- Prioritize findings over rewrite suggestions

### `adr-reconciler`

- Compare ADR claims with the current state of `specs` / `src` / `tests`
- Return an evidence matrix across `adr`, `specs`, `src`, and `tests` for each finding
- Use the finding types `policy-drift`, `missing-adr-update`, `missing-evidence`, and `stale-reference`
- Return `severity` (`high` / `medium` / `low`) and `recommended action` for each finding
- Do not auto-edit ADR text

### `adr-reviewer`

- Perform design review of ADR drafts
- Assume there are no unresolved `adr-linter` `high` / `medium` findings in the current draft
- Include related `docs/` in scope when the ADR touches public API / CLI / MCP
- Use the finding types `decision-gap`, `scope-conflict`, `evidence-risk`, `rollout-gap`, and `ownership-escalation`
- Return verdicts `ready`, `revise`, and `escalate`
- Do not repeat structural checks already handled by `adr-linter`; focus on design issues, conflicts with existing ADRs, and compatibility / rollout / fallback / safety impact
- Escalate issues outside AI ownership back to humans

### `adr-indexer`

- Scan existing ADRs and their metadata, then synchronize `README.md`, `index.yaml`, and `decision-map.md`
- Return findings when status, domain, supersede relationships, or related specs are inconsistent
- Treat index artifacts as derived views of the ADR source text, not as the source of truth

## Pre-merge checks

- ADR conclusions do not conflict with specs
- Contracts written in specs are backed by tests
- If an ADR supersedes an existing ADR, the cross-references are filled in
- Even when an ADR is unnecessary, the reason is left in the issue or PR
- Even when an ADR is unnecessary, the `specs`, `src`, and `tests` evidence used for the decision remains traceable
- `adr-reconciler` `high` findings are not left unresolved at merge time
- `adr-linter` `high` / `medium` findings are not left unresolved at merge time
- `adr-reviewer` `revise` verdicts or `high` / `medium` findings are not left unresolved at merge time
- Issues returned as `escalate` by `adr-reviewer` are not left unresolved at merge time

## Post-merge / periodic audit checks

- If `adr-reconciler` returns `high` findings, leave the target ADR and the drifting spec / test / code path in the issue or PR
- If the drift indicates a policy-level change, go back to `adr-suggester` and re-evaluate `required` / `recommended` / `not-needed`
- If an ADR is added, updated, or superseded, update the derived artifacts with `adr-indexer`
- `index.yaml` and `decision-map.md` must not disagree with the ADR source text on status or supersede relationships

## Future phases

- More detailed review automation and PR-bot integration
