---
name: adr-drafter
description: Draft a new ExStruct ADR or propose an update to an existing ADR from an issue, PR, diff, tests, and specs. Use when an ADR is required or recommended and you need a structured draft with context, decision, consequences, and evidence.
---

# ADR Drafter

Draft a policy record, not an implementation memo.

## Read

1. `dev-docs/agents/adr-governance.md`
2. `dev-docs/agents/adr-criteria.md`
3. `dev-docs/agents/adr-workflow.md`
4. `dev-docs/adr/template.md`
5. Related ADRs under `dev-docs/adr/`

## Workflow

1. Confirm whether the target is a new ADR or an update to an existing ADR.
2. Read the issue, diff, relevant specs, tests, and implementation files.
3. Extract:
   - policy-level context
   - chosen decision
   - positive and negative consequences
   - concrete evidence from `tests`, `code`, and `related specs`
4. Produce one of these outputs:
   - a new ADR draft
   - an update proposal for an existing ADR
   - `ADR not needed` with rationale

## Draft Rules

- Keep `why` in the ADR and leave detailed `how` in code or specs.
- Fill `Tests`, `Code`, and `Related specs` whenever possible.
- If the policy replaces an old one, state the supersede relation explicitly.
- If evidence is missing, say so instead of inventing support.
