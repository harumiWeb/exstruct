# Safe Editing

Use this file for risky edits, unfamiliar workbooks, or any request where a
wrong apply would be expensive.

## Standard flow

1. Determine whether the request creates a new workbook or edits an existing
   one.
2. Validate the input workbook when readability is uncertain for an existing
   workbook edit.
3. Build or inspect the patch-op JSON.
4. Run `exstruct patch --dry-run` for risky edits to an existing workbook.
5. Run `exstruct make --dry-run` for risky new-workbook requests when the
   chosen ops and output format support dry-run.
6. Inspect `PatchResult`, warnings, `patch_diff`, and `engine`.
7. Re-run without `--dry-run` only after the plan is acceptable.
8. Perform lightweight verification or deeper verification based on risk.

## Ambiguity rules

- Ask for clarification before apply when sheet names, target cells/ranges,
  output naming, or overwrite policy are unclear.
- If the user only states a goal, inspect the op schema before inventing a
  payload shape.

## Risk checks

- Treat destructive or broad edits as risky:
  - overwrite/rename conflict decisions
  - merge / unmerge operations
  - range-wide formula fills
  - style changes across large regions
  - chart creation or COM-only workflows
- `create_chart` requests and `.xls` create/edit workflows do not support
  `--dry-run`, `--return-inverse-ops`, or `--preflight-formula-check`;
  inspect ops, run only supported validation, and explain the backend
  constraint before any non-dry-run apply.
- Prefer `--backend openpyxl` when you want the dry run and the later apply run
  to use the same engine.
- Explain that `--backend auto` can resolve to openpyxl for a dry-run request
  and COM for a later non-dry-run request on COM-capable hosts.

## Failure behavior

- Surface schema or validation failures directly.
- Do not silently retry with a different conceptual operation.
- If the request depends on unsupported runtime features, explain the
  constraint and route to the correct backend or MCP workflow.
