# ADR-0003: Output Serialization Omission Policy

## Status

`accepted`

## Background

ExStruct is used for LLM / RAG preprocessing, making output size important.
Backend metadata is useful for debugging and fidelity analysis, but always including it by default increases token cost during normal use.

## Decision

- Backend metadata such as `provenance`, `approximation_level`, and `confidence` is omitted by default in serialized output.
- Backend metadata remains accessible via an explicit opt-in flag in the public interface.
- This omission policy is part of the serialization contract and must be consistent across API, CLI, MCP, and schema expectations.

## Consequences

- New metadata fields should be omitted by default unless there is a clear public need.
- The public docs must describe how to opt in.
- Tests must cover both the default-omit path and the explicit-include path.

## Rationale

- Tests: `tests/models/test_models_export.py`, `tests/models/test_schemas_generated.py`
- Code: `src/exstruct/io/serialize.py`, `src/exstruct/models/__init__.py`
- Related specs: `docs/api.md`, `docs/cli.md`

## Supersedes

- None

## Superseded by

- None
