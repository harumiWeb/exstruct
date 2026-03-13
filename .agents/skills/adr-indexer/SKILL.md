---
name: adr-indexer
description: Maintain ExStruct ADR index artifacts by scanning ADR files and synchronizing README, index.yaml, and decision-map.md metadata. Use when ADRs are added, updated, superseded, or reclassified and the repository needs a refreshed ADR map for humans and AI agents.
---

# ADR Indexer

Refresh the derived ADR map after the source ADRs change.

## Read

1. `dev-docs/agents/adr-governance.md`
2. `dev-docs/agents/adr-workflow.md`
3. `dev-docs/specs/adr-index.md`
4. All ADR files under `dev-docs/adr/`
5. Existing `dev-docs/adr/README.md`, `dev-docs/adr/index.yaml`, `dev-docs/adr/decision-map.md`

## Workflow

1. Scan the ADR files and collect normalized metadata:
   - `id`
   - `title`
   - `status`
   - `primary_domain`
   - `domains`
   - `supersedes`
   - `superseded_by`
   - `related_specs`
2. Compare the normalized metadata with `README.md`, `index.yaml`, and `decision-map.md`.
3. Update the derived artifacts so that status, primary domain, domain grouping, and supersede relationships match the ADR source files.
4. Report any ambiguity that cannot be derived from the ADR text alone.

## Output Contract

Return a concise summary with:

- `updated artifacts`
- `added or changed ADR entries`
- `consistency findings`

Treat `ADR-xxxx-*.md` as the source of truth. Do not invent domains, statuses, or supersede links that the ADR text cannot justify.
