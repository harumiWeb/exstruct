# TODO: MCP `create_chart` 追加

## Plan

- [x] `tasks/feature_spec.md` に仕様確定版を記載
- [x] `tasks/todo.md` に実装計画を記載
- [ ] `PatchOpType` / `PatchValueKind` に `create_chart` / `chart` を追加
- [ ] `PatchOp` モデルに `create_chart` 用フィールドとvalidatorを追加
- [ ] `specs.py` / `op_schema.py` / `list_ops` / `describe_op` を更新
- [ ] `service.py` に COM専用ガードを追加
- [ ] `internal.py` openpyxl側に `create_chart` 明示拒否を追加
- [ ] `internal.py` xlwings側に `_apply_xlwings_create_chart` を追加
- [ ] `server.py` の patch tool docstring を更新
- [ ] `docs/mcp.md` / `README.md` / `README.ja.md` / `CHANGELOG.md` を更新
- [ ] テスト追加・更新
- [ ] `uv run pytest -m "not com and not render"` を実行
- [ ] `uv run task precommit-run` を実行

## Review

- Summary:
- Risks:
- Follow-ups:
