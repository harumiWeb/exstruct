---
name: adr-linter
description: Review an ExStruct ADR draft for required sections, status values, evidence quality, supersede links, and balanced consequences. Use when an ADR already exists in draft form and you need findings before review or merge.
---

# ADR Linter

Review the ADR as a decision record, not as prose polish.

## Read

1. `dev-docs/agents/adr-governance.md`
2. `dev-docs/agents/adr-workflow.md`
3. `dev-docs/adr/template.md`
4. Related ADRs if the draft mentions supersede or overlap

## Workflow

1. Validate the status value.
2. Check that the ADR has `状態`, `背景`, `決定`, `影響`, `根拠`, `Supersedes`, and `Superseded by`.
3. Verify that `根拠` contains concrete `Tests`, `Code`, or `Related specs`.
4. Check that consequences include tradeoffs, not only benefits.
5. If supersede or replacement is claimed, verify the referenced ADR links are present and consistent.

## Output Contract

Return findings first, ordered by severity.

- `high`: contract hole, missing decision, invalid status, missing supersede linkage
- `medium`: weak context, weak evidence, one-sided consequences
- `low`: clarity or consistency issues

If no findings exist, say that explicitly and mention any residual review risk.
