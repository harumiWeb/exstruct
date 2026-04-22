# Contributing guide for AI agents

This file contains **special guidelines** for AI coding agents such as ChatGPT, Cursor, and Copilot.

## Principles

1. Read `docs/` as the public contract, `dev-docs/specs/` as the internal specification, and `dev-docs/adr/` as the record of decision rationale.
2. Do not write code that contradicts model definitions in `dev-docs/specs/data-model.md`.
3. The `core` layer is for extraction only. Integration logic is centralized in `modeling.py`, and `integrate.py` stays a thin entry point for pipeline invocation.
4. The `models` layer must remain completely side-effect-free.
5. Do not mix I/O processing with core logic.
6. Keep exception handling fail-safe.
7. Update the roadmap whenever you add a new feature.

## Reference priority

1. `docs/`
2. `dev-docs/specs/`
3. `dev-docs/adr/`
4. `tests/`
5. `src/`

Responsibility split:

- ADR = why a decision was made
- specs = what is guaranteed
- tests = evidence of the behavior
- src = how it is implemented

## Task separation for AI

- New extraction features or semantic-analysis algorithms -> `core/`
- New data structures -> `models/`
- New output formats -> `io/`
- CLI features -> `cli/`

## Coding guidelines

Always follow these rules:

- Add type hints to every argument and return value
- Keep one function to one responsibility
- Return `BaseModel` at boundaries and dataclasses internally
- Keep imports in the correct order
- Write docstrings in Google style
- Split functions before they become too complex
- Return Pydantic models rather than JSON blobs or dictionaries

## Testing policy

- Use `pytest` and `pytest-mock` as the testing framework
- Place sample Excel files in `/tests/data/*.xlsx`
- Prefer regression tests that lock down Pydantic/dataclass model agreement
- Use Ruff and mypy for static analysis, and write implementations that pass both
