# ADR-0002: Rich Backend Fallback Policy

## Status

`accepted`

## Background

Rich extraction depends on runtimes that can be absent or unstable, such as Excel COM on Windows and LibreOffice in non-COM environments.
A consistent fallback contract is needed to keep extraction results useful while not hiding the reason that rich artifacts are missing.

## Decision

- Runtime unavailability is treated as a normal fallback condition, not an exceptional product failure.
- Fallback reasons are recorded and logged explicitly through `FallbackReason`.
- Even when the rich backend fails, ExStruct does not discard the entire extraction result; it retains the best safe result available.
- Although COM and LibreOffice differ internally, the high-level fallback policy is kept consistent across both.

## Consequences

- When introducing a new backend, its fallback behavior must be defined upfront.
- When changing error handling, regression tests are required that verify both the return data shape and the fallback reason.
- Internal runtime hardening must not silently change the public fallback contract.

## Rationale

- Tests: `tests/integration/test_integrate_fallback.py`, `tests/utils/test_logging_utils.py`
- Code: `src/exstruct/core/pipeline.py`, `src/exstruct/errors.py`
- Related specs: `dev-docs/specs/excel-extraction.md`, `docs/concept.md`

## Supersedes

- None

## Superseded by

- None
