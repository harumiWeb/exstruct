# ADR-0005: PathPolicy Safety Boundary

## Status

`accepted`

## Background

MCP tools manipulate file paths provided by external clients.
A clear safety boundary is needed regarding which paths are allowed, how output paths are resolved, and where path validation is placed.

## Decision

- `PathPolicy` is the central boundary for validating MCP file access and output location rules.
- Callers normalize and validate paths through `PathPolicy`-aware helpers rather than duplicating path checks at each tool call site.
- Safety checks are treated as part of the tool execution contract, not optional hygiene.

## Consequences

- New MCP file-manipulating tools must integrate with `PathPolicy`.
- Direct filesystem access in MCP handlers should be treated with suspicion unless justified by a tested boundary helper.
- Changing path behavior requires tests covering allow, deny, and normalization cases.

## Rationale

- Tests: `tests/mcp/test_path_policy.py`, `tests/mcp/test_validate_input.py`, `tests/mcp/shared/test_output_path.py`
- Code: `src/exstruct/mcp/io.py`, `src/exstruct/mcp/server.py`, `src/exstruct/mcp/render_runner.py`
- Related specs: `docs/mcp.md`

## Supersedes

- None

## Superseded by

- None
