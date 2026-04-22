# ADR-0001: Extraction Mode Responsibility Boundaries

## Status

`superseded`

## Background

ExStruct provides multiple extraction modes, each with different guarantees and runtime requirements.
Without an explicit decision record, the responsibilities of `light`, `libreoffice`, `standard`, and `verbose` are easy to conflate, and the boundaries tend to erode when adding new artifacts or validation rules.

## Decision

- `light` maintains a minimal configuration and does not introduce any rich runtime dependency.
- `libreoffice` is a non-COM best-effort rich mode for `.xlsx/.xlsm` and explicitly rejects PDF, image, and auto page-break export.
- `standard` and `verbose` are maintained as COM-enabled modes that perform higher-fidelity native Excel extraction.
- Per-mode validation is treated as part of the product contract, not merely an implementation detail.

## Consequences

- When adding a new feature, the mode responsible for that behavior must be stated explicitly.
- Validation logic must be consistent across the API, CLI, and engine entry points.
- The responsibility table per mode is maintained explicitly rather than relying solely on the implicit knowledge embedded in tests.

## Rationale

- Tests: `tests/test_constraints.py`
- Code: `src/exstruct/constraints.py`, `src/exstruct/engine.py`
- Related specs: `docs/api.md`, `docs/cli.md`, `dev-docs/specs/excel-extraction.md`

## Supersedes

- None

## Superseded by

- `ADR-0010`
