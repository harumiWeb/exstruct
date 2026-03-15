# ExStruct Developer Documentation

`dev-docs/` contains internal documentation for maintainers and AI agents.

## Purpose

- `docs/`: user-facing documentation published with MkDocs
- `dev-docs/`: internal documentation for implementation, maintenance, and design decisions

## Reading order

When you change the codebase, read these materials in the following order:

1. Check the public contract in `docs/`.
2. Check the current internal specifications in `dev-docs/specs/`.
3. Check design decisions and constraints in `dev-docs/adr/`.
4. Check behavior evidence in `tests/`.
5. Check implementation details in `src/`.

## Responsibility split

- ADR = why a decision was made
- specs = what is guaranteed
- tests = evidence that the behavior exists
- src = how it is implemented

## Directory structure

- `adr/`: design decisions and trade-offs
- `agents/`: operating guides for AI agents
- `architecture/`: implementation structure and extension guides
- `specs/`: current internal specifications
- `testing/`: test requirements and validation policy
