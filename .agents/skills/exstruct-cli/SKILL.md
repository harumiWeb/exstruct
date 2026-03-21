---
name: exstruct-cli
description: Use ExStruct CLI to validate, inspect, create, and edit Excel workbooks safely. Trigger when an agent needs `exstruct patch`, `exstruct make`, `exstruct validate`, `exstruct ops list`, or `exstruct ops describe`, especially for create-vs-edit decisions, dry-run workflows, backend constraints, or safe workbook-edit guidance.
---

# ExStruct CLI

Use the existing editing CLI as the default local operational interface for
ExStruct workbook create/edit requests.

## Select a command

- Use `exstruct make` when the user wants a new workbook created and populated.
- Use `exstruct patch` when the user wants an existing workbook edited.
- Use `exstruct validate` when workbook readability is uncertain or before a
  risky edit on an unfamiliar file.
- Use `exstruct ops list` when the required operation is still unclear.
- Use `exstruct ops describe <op>` when you know the likely op name but need
  the exact schema or constraints.
- Hand the workflow off to MCP when the user needs host-owned path policy,
  transport mapping, artifact mirroring, or other server-managed behavior.

## Safety rules

- Do not invent unsupported patch ops.
- Do not apply destructive edits immediately when the request is ambiguous.
- Prefer `--dry-run` before risky edits.
- Pin `--backend openpyxl` when the dry run and the real apply must use the
  same engine.
- Explain backend-specific failures directly instead of hiding them behind a
  generic retry.

## Workflow

1. Decide whether the request is create, edit, or actually an MCP-hosted task.
2. Identify the needed op or inspect the public op schema.
3. Validate the workbook when the file, backend, or request is uncertain.
4. Run `--dry-run` for risky edits and inspect `PatchResult`.
5. Apply the real edit only after the planned change is acceptable.
6. Verify the result with the lightest step that still matches the user's risk.

## References

- `references/command-selection.md`
- `references/safe-editing.md`
- `references/ops-guidance.md`
- `references/verify-workflows.md`
- `references/backend-constraints.md`
