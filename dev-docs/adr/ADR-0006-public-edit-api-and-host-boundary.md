# ADR-0006: Public Edit API and Host-Owned Safety Boundary

## Status

`accepted`

## Background

ExStruct already had a capable workbook editing pipeline, but it was reachable
primarily through MCP-facing modules such as `exstruct.mcp.patch_runner` and the
`exstruct_patch` / `exstruct_make` tool handlers. That made normal Python usage
awkward and blurred the boundary between editing behavior and host-owned safety
policy.

The issue is not only discoverability. The MCP layer also owns path sandboxing,
artifact mirroring, and server execution concerns that should not be mandatory
for ordinary library callers. At the same time, existing patch models, backend
selection rules, and warning/error payloads are already tested and should remain
stable while the public surface is promoted.

## Decision

- `exstruct.edit` is adopted as the first-class public Python API for workbook
  editing.
- Phase 1 public entry points are `patch_workbook(PatchRequest)` and
  `make_workbook(MakeRequest)`.
- The existing patch contract is preserved for Phase 1:
  - `PatchOp`, `PatchRequest`, `MakeRequest`, `PatchResult`
  - op names
  - normalization behavior
  - schema-discovery metadata
  - backend warning/error payload shape
- MCP remains a host/integration layer. It continues to own:
  - `PathPolicy`
  - MCP tool input/output mapping
  - artifact mirroring
  - server defaults, thread offloading, and runtime controls
- Phase 1 may reuse the proven `exstruct.mcp.patch.*` execution pipeline under
  the hood while `exstruct.edit` becomes the canonical public import path.

## Consequences

- Python callers now have a direct, library-oriented entry point that does not
  require MCP-specific path restrictions.
- MCP compatibility remains intact because the old import paths stay available.
- Operation schemas and alias normalization can now be treated as part of the
  public editing surface, not only MCP documentation.
- The transition keeps two module trees in play during Phase 1, which is less
  clean than a full implementation relocation but materially reduces risk while
  the new public surface is established.
- Future phases can move more execution internals under `exstruct.edit` without
  reopening the public contract question.

## Rationale

- Tests:
  - `tests/edit/test_api.py`
  - `tests/mcp/patch/test_normalize.py`
  - `tests/mcp/test_patch_runner.py`
  - `tests/mcp/test_make_runner.py`
  - `tests/mcp/patch/test_service.py`
  - `tests/mcp/test_tools_handlers.py`
- Code:
  - `src/exstruct/edit/__init__.py`
  - `src/exstruct/edit/api.py`
  - `src/exstruct/edit/service.py`
  - `src/exstruct/mcp/patch_runner.py`
  - `src/exstruct/mcp/tools.py`
- Related specs:
  - `dev-docs/specs/editing-api.md`
  - `dev-docs/specs/data-model.md`
  - `docs/api.md`
  - `docs/mcp.md`

## Supersedes

- None

## Superseded by

- None
