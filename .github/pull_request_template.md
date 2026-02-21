## Summary

- [ ] Describe the scope and motivation.
- [ ] Link related issue/spec (if any).

## Acceptance Criteria (MCP UX Hardening)

- [ ] AC-01: 入力互換（`name`, `col/row`, `width/height`）
- [ ] AC-02: HEX自由指定（6/8桁, `#`有無）
- [ ] AC-03: 色概念分離（`color` と `fill_color`）
- [ ] AC-04: `draw_grid_border` shorthand（`range`）
- [ ] AC-05: path UX（root基準解決と診断）
- [ ] AC-06: 後方互換（既存成功ケース維持）

## Validation

- [ ] `uv run task precommit-run`
- [ ] Added/updated tests for changed behavior.
- [ ] Updated docs (`docs/mcp.md`, `docs/README.ja.md`, release notes).
