# Editing CLI Specification

This document defines the Phase 2 public editing CLI contract.

## Command surface

- Editing commands are exposed from the existing `exstruct` console script.
- Phase 2 commands:
  - `exstruct patch`
  - `exstruct make`
  - `exstruct ops list`
  - `exstruct ops describe`
  - `exstruct validate`
- The legacy extraction entrypoint `exstruct INPUT.xlsx ...` remains valid and
  is not rewritten to `exstruct extract` in Phase 2.

## Canonical usage documentation obligations

- Public docs must describe the editing CLI as the canonical operational /
  agent interface for workbook editing.
- Public docs should recommend the `dry_run -> inspect PatchResult -> apply`
  workflow for edit operations, but must qualify that `backend="auto"` can
  use openpyxl for the dry run and COM for the real apply on COM-capable
  hosts; when same-engine comparison matters, docs should tell users to pin
  `backend="openpyxl"`.
- Public docs must distinguish the local CLI from:
  - `exstruct.edit` for embedded Python usage
  - MCP for host-owned path policy, transport, and artifact behavior

## Dispatch and compatibility rules

- `exstruct.cli.main` dispatches to the editing parser only when the first
  token is one of the Phase 2 editing subcommands.
- All other invocations continue to use the extraction parser and
  `process_excel` path unchanged.
- Phase 2 does not add a new console script or top-level Python export.

## Patch and make commands

- `patch` is the CLI wrapper over `exstruct.edit.patch_workbook`.
- `make` is the CLI wrapper over `exstruct.edit.make_workbook`.
- Shared request flags:
  - `--sheet`
  - `--on-conflict {overwrite,skip,rename}`
  - `--backend {auto,com,openpyxl}`
  - `--auto-formula`
  - `--dry-run`
  - `--return-inverse-ops`
  - `--preflight-formula-check`
  - `--pretty`
- `patch` requires:
  - `--input PATH`
  - `--ops FILE|-`
- `patch` optionally accepts `--output PATH`; when omitted, the existing patch
  output defaulting behavior remains in effect.
- `make` requires `--output PATH`.
- `make` accepts optional `--ops FILE|-`; when omitted, `ops=[]`.

## Ops input contract

- `--ops` reads UTF-8 JSON from a file path or stdin marker `-`.
- The top-level JSON value must be an array.
- Each array item follows the existing public patch-op normalization rules
  exposed from `exstruct.edit`, including alias normalization and JSON-string
  op coercion.
- Phase 2 does not accept a request-envelope JSON document on the CLI.

## Output contract

- `patch` and `make` serialize `PatchResult` to stdout as JSON.
- `validate` serializes the existing input validation result shape:
  - `is_readable`
  - `warnings`
  - `errors`
- `ops list` returns compact summaries with `op` and `description`.
- `ops describe` returns detailed patch-op schema metadata for one op.
- `--pretty` applies `indent=2` JSON formatting to all Phase 2 editing
  commands.

## Exit-code rules

- `patch` / `make` exit `0` when the serialized `PatchResult` has
  `error is None`; otherwise they exit `1`.
- `validate` exits `0` when `is_readable=true`; otherwise `1`.
- `ops list` exits `0` on success.
- `ops describe` exits `1` for unknown op names.
- JSON parse failures, request validation failures, and local I/O failures are
  reported as stderr CLI errors and exit `1`; Phase 2 does not introduce a
  separate generic JSON error envelope for these cases.

## Explicit non-goals for Phase 2

- No `exstruct extract` subcommand
- No backup / confirmation / allow-root / deny-glob flags
- No summary-mode output
- No changes to backend selection or fallback policy
- No changes to MCP tool contracts
- No new public Python validation API
