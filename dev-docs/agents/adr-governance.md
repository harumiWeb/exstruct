# ADR governance

This document defines when AI agents should create or update ADRs in ExStruct, and what counts as evidence.

## Purpose

- Prevent design decisions that affect long-term maintenance from being buried in issue threads or diffs
- Preserve the responsibility split with `dev-docs/specs/`, `tests/`, and `src/`
- Make the relationship between existing ADRs and new changes explicit, so decision rationale does not go stale

## Role of each document

- ADR = why a policy was chosen
- specs = what is guaranteed
- tests = evidence that the guarantee exists
- src = how it is implemented

Do not put implementation notes alone into ADRs. Simple procedures and temporary work order belong in `tasks/`.

## When to create a new ADR

Prefer a new ADR when any of the following apply:

- adding a new policy or responsibility boundary
- introducing a decision on a different issue that existing ADRs cannot explain
- comparing multiple options or accepting a long-term trade-off
- addressing a question that may recur in the future

## When to update an existing ADR

Prefer updating an existing ADR when all of the following are true:

- the existing ADR's policy remains intact, but its context, evidence, or related specs should be strengthened
- the same issue covered by the existing ADR gains additional evidence or clarifying constraints
- implementation or documentation cleanup only requires references to be refreshed

If the conclusion of the ADR itself changes, consider a new ADR plus `supersede` instead of an update.

## Handling supersedes

- If the old ADR no longer has authority, create a new ADR and mark the old ADR as `superseded`
- Update both `Superseded by` on the old ADR and `Supersedes` on the new ADR
- Do not do partial overwrite without stating which policy was replaced

## Evidence requirements

An ADR must have at least one piece of concrete evidence.

- Tests: regression tests, contract tests, integration tests
- Code: key files that implement the decision
- Related specs: internal specs or public docs that define the current contract

The recommended baseline is to fill all three categories: `Tests`, `Code`, and `Related specs`.

## Draft review rules

When handling a `proposed` ADR before merge, separate structural checks from design review.

- `adr-linter` checks status, required sections, presence of evidence, and consistency of supersede links
- `adr-reviewer` is used only when the current draft has no unresolved `adr-linter` `high` / `medium` findings. For ADRs that touch public API / CLI / MCP, include the related `docs/` in scope and review decision quality, conflicts with existing ADRs / specs, evidence strength, and compatibility / rollout / fallback / safety impact
- `adr-reviewer` uses the verdicts `ready`, `revise`, and `escalate`
- If the verdict is `revise`, update the draft and rerun `adr-linter` and `adr-reviewer`
- Use `escalate` when the issue includes matters outside AI ownership, such as public API break judgment, security / license decisions, large directory restructures, or unresolved product / spec policy
- `ready` does not mean the merge itself is approved. Questions that require human final judgment may still remain

## Reconciliation rules

When an ADR is `proposed` or `accepted`, check for drift with `adr-reconciler` during changes or periodic reviews.

- Audit at the claim level and gather an evidence matrix across `adr`, `specs`, `src`, and `tests`
- Use at least the following finding types:
  - `policy-drift`
  - `missing-adr-update`
  - `missing-evidence`
  - `stale-reference`
- Findings carry `severity` values of `high`, `medium`, or `low`
- Resolve `high` findings before merge or create an explicit follow-up
- Use the following `recommended action` values:
  - `update-adr`
  - `new-adr`
  - `update-specs`
  - `add-tests`
  - `no-action`

`adr-reconciler` only returns audit results. It does not auto-edit ADR or spec text. If a policy-level change is suspected, return to `adr-suggester` and `adr-drafter`.

## Handling index artifacts

Treat the following files as derived artifacts from ADR source text:

- `dev-docs/adr/README.md`
- `dev-docs/adr/index.yaml`
- `dev-docs/adr/decision-map.md`

Update them when a new ADR is added, when status changes, when supersede relationships change, when domain classification changes, or when related specs change.

## Status rules

The only allowed status values are:

- `proposed`
- `accepted`
- `superseded`
- `deprecated`

Do not use `draft` or custom statuses. Even at the drafting stage, the ADR document itself should use `proposed`.

## Required AI-agent behavior

- Check existing ADRs that may be related before starting the change
- For ADR-worthy issues, do not mix up `why` and `what`
- When proposing a new ADR, list the related `tests`, `code`, and `specs`
- Even when deciding an ADR is unnecessary, leave a short note explaining why it is not a policy-level change
- When creating or updating an ADR, run `adr-reconciler` as needed and do not leave `high` findings unresolved
- When adding, updating, or superseding an ADR, also synchronize the derived index artifacts
- When reviewing an ADR draft, clear the current draft's `adr-linter` `high` / `medium` findings first, and if the ADR touches a public surface, review the related `docs/` as well before using `adr-reviewer` to capture only design-level findings

## Non-goals

- Do not define the public contract with ADRs alone
- Do not prove implementation correctness with ADRs alone
- Do not use ADRs as detailed design docs for CI or bot integrations
