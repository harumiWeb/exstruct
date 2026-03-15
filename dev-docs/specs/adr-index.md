# ADR Index Specification

This document defines the internal contract for the derived artifacts maintained by `adr-indexer`.

## Target Artifacts

- `dev-docs/adr/README.md`
- `dev-docs/adr/index.yaml`
- `dev-docs/adr/decision-map.md`

The source of truth is each `ADR-xxxx-*.md` body; the index artifacts are derived from it.

## `index.yaml` Contract

- The top-level key is `adrs`
- Each entry must have at least the following:
  - `id`
  - `title`
  - `status`
  - `path`
  - `primary_domain`
  - `domains`
  - `supersedes`
  - `superseded_by`
  - `related_specs`
- `status` allows only `accepted`, `proposed`, `superseded`, `deprecated`
- `primary_domain` must be one of the values in `domains`
- `primary_domain` is the source of truth for the "Primary Domain" column in README
- `domains` is limited to short classification tags for AI agents to traverse related ADRs
- `supersedes` and `superseded_by` are arrays of ADR IDs and must not exist on only one side

## `decision-map.md` Contract

- Group ADRs by domain
- Headings correspond one-to-one with each element of the `domains` array; do not combine multiple domains under one heading
- Each ADR must include at least its ID, title, and status
- When a supersede relationship exists, make it explicit on the decision map as well
- Do not create a heading for a domain that has no ADRs

## `README.md` Contract

- Contains a short human-readable list
- status and primary domain must match the `status` / `primary_domain` in `index.yaml`
- Delegate detailed relationship graphs and machine-readable metadata to `decision-map.md` and `index.yaml`

## Update Triggers

Update index artifacts when any of the following occurs.

- New ADR added
- ADR status changed
- `Supersedes` / `Superseded by` changed
- Related spec referenced by the ADR changed
- Domain classification added or revised
