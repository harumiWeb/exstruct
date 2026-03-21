# Command Selection

Use this file when the user intent is clear enough to route to one CLI command
but the exact command is still undecided.

## Primary routing

- New workbook output requested: use `exstruct make`.
- Existing workbook path provided and the request changes it: use
  `exstruct patch`.
- Readability or file-health check requested: use `exstruct validate`.
- Unknown capability or op name: use `exstruct ops list` first, then
  `exstruct ops describe <op>`.
- Host-owned path restrictions, transport, approval-aware execution, or
  artifact mirroring requested: use MCP guidance instead of this Skill.

## Request patterns

- "Create a workbook with these starter sheets and values."
  - `make`
- "Edit this workbook and set/update/merge/style cells."
  - `patch`
- "I do not know which op supports this change."
  - `ops list` or `ops describe`
- "Check whether this workbook is readable before we touch it."
  - `validate`
- "Do this through the MCP server / Claude Desktop / host-managed policy."
  - MCP docs, not the local CLI Skill

## Escalate instead of guessing

- Clarify the request before applying edits when the sheet, range, output path,
  or intended effect is underspecified.
- If the user requests behavior outside the published op schema, say so and
  offer the nearest supported workflow or a manual alternative.
