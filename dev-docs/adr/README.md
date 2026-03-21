# ADR

ADRs record what was decided, under which constraints, and which trade-offs were accepted.

## Purpose

- Record decisions that affect structure and long-term maintenance
- Preserve alternatives not chosen and non-goals
- Enable tracking of which ADR superseded which in the future

## Status

- `proposed`
- `accepted`
- `superseded`
- `deprecated`

## Numbering

- Use `ADR-0001`, `ADR-0002`, ... in the order they are added to the repository
- Numbers are fixed even if the title changes

## Relationship to Other Documents

- ADR explains "why it was done this way"
- `dev-docs/specs/` explains "what is guaranteed"
- `tests/` provides evidence that the behavior actually exists

## Index Artifacts

- `index.yaml`: machine-readable ADR metadata index for AI agents
- `decision-map.md`: overview of related ADRs and supersede relationships by domain

## Index

| ID | Title | Status | Primary Domain |
| --- | --- | --- | --- |
| `ADR-0001` | Extraction Mode Responsibility Boundaries | `accepted` | `extraction` |
| `ADR-0002` | Rich Backend Fallback Policy | `accepted` | `backend` |
| `ADR-0003` | Output Serialization Omission Policy | `accepted` | `schema` |
| `ADR-0004` | Patch Backend Selection Policy | `accepted` | `mcp` |
| `ADR-0005` | PathPolicy Safety Boundary | `accepted` | `safety` |
| `ADR-0006` | Public Edit API and Host-Owned Safety Boundary | `accepted` | `editing` |
| `ADR-0007` | Editing CLI as Public Operational Interface | `accepted` | `editing` |
| `ADR-0008` | Extraction CLI Runtime Capability Validation | `accepted` | `cli` |
| `ADR-0009` | Single CLI Skill for Agent Workflows | `proposed` | `agents` |
