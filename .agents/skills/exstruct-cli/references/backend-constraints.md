# Backend Constraints

Use this file when the workflow depends on backend selection or runtime
capabilities.

## Backend summary

- `openpyxl`
  - recommended backend to pin for ordinary workbook edits, especially when
    you want the same engine for dry-run and apply
  - supports `--dry-run`, `--return-inverse-ops`, and
    `--preflight-formula-check`
- `com`
  - required for COM-only editing behavior such as `create_chart`
  - required for `.xls` creation/editing
- `auto`
  - follows ExStruct's existing backend-selection policy
  - can resolve to openpyxl for a dry-run request and COM for a later apply
    request

## Constraints to call out

- `create_chart` is COM-only and does not support `--dry-run`,
  `--return-inverse-ops`, or `--preflight-formula-check`.
- `.xls` workflows require COM, are not valid with `backend=openpyxl`, and do
  not support `--dry-run`, `--return-inverse-ops`, or
  `--preflight-formula-check`.
- Requests that need host-owned path restrictions, transport mapping, or
  artifact mirroring belong to MCP, not the local CLI Skill.

## Failure behavior

- State clearly when the requested backend or capability is unsupported.
- Do not pretend that `auto` guarantees the same engine across separate dry-run
  and apply requests.
- Redirect to MCP only for host-policy concerns, not as a generic substitute
  for unsupported CLI behavior.
