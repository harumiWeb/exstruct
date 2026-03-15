# ADR-0004: Patch Backend Selection Policy

## Status

`accepted`

## Background

`exstruct_patch` supports multiple backends with different capabilities and safety constraints.
Backend selection directly affects compatibility, fallback behavior, and allowed operations, so an explicit decision record is needed rather than ad-hoc rules at each call site.

## Decision

- `backend="auto"` prefers COM when available and falls back to openpyxl only for permitted runtime failures.
- `backend="com"` is an explicit selection and does not silently fall back to openpyxl.
- `backend="openpyxl"` is maintained as a safe pure-Python path; feature gaps such as chart creation and `.xls` handling are made explicit.
- Capability restrictions are enforced as request validation where they can be determined upfront, rather than deferring to backend execution time.

## Consequences

- When adding backend-specific features, the behavior for each of `auto`, `com`, and `openpyxl` must be stated explicitly.
- Changing the fallback policy requires adding tests covering both positive and negative cases.
- The documentation's backend capability table must be kept in sync with runtime validation.

## Rationale

- Tests: `tests/mcp/patch/test_service.py`, `tests/mcp/patch/test_models_internal_coverage.py`
- Code: `src/exstruct/mcp/patch/runtime.py`, `src/exstruct/mcp/patch/service.py`, `src/exstruct/mcp/server.py`
- Related specs: `docs/mcp.md`, `dev-docs/specs/patch/legacy-dependency-inventory.md`

## Supersedes

- None

## Superseded by

- None
