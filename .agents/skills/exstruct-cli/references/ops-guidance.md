# Ops Guidance

Use this file when the user knows the desired workbook outcome but the exact
patch op is unclear.

## How to inspect ops

- Run `exstruct ops list` when you need the supported op names and one-line
  descriptions.
- Run `exstruct ops describe <op>` when you need exact fields, aliases,
  backend constraints, or examples.

## Common patterns

- Single cell value/formula updates:
  - inspect `set_value`, `set_formula`, and the conditional variants
- New worksheet creation:
  - inspect `add_sheet`
- Rectangular data loads:
  - inspect `set_range_values`
- Styling and layout:
  - inspect `set_style`, `set_alignment`, `set_fill_color`, `set_bold`,
    `auto_fit_columns`, `set_dimensions`
- Table or chart workflows:
  - inspect `apply_table_style` and `create_chart`

## Rules

- Never invent an op name that `ops list` does not expose.
- Never guess required fields when `ops describe` can answer the question.
- Treat `create_chart` and `.xls` workflows as backend-sensitive and confirm the
  backend before apply.
- If the user asks for behavior outside the op schema, say it is unsupported
  and offer the closest supported path.
