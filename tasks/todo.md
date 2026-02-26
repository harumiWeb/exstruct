## Plan

- [x] `tasks/feature_spec.md` に仕様確定版を記載
- [x] `tasks/todo.md` に実装計画を記載
- [x] `PatchOpType` / `PatchValueKind` に `create_chart` / `chart` を追加
- [x] `PatchOp` モデルに `create_chart` 用フィールドとvalidatorを追加
- [x] `specs.py` / `op_schema.py` / `list_ops` / `describe_op` を更新
- [x] `service.py` に COM専用ガードを追加
- [x] `internal.py` openpyxl側に `create_chart` 明示拒否を追加
- [x] `internal.py` xlwings側に `_apply_xlwings_create_chart` を追加
- [x] `server.py` の patch tool docstring を更新
- [x] `docs/mcp.md` / `README.md` / `README.ja.md` / `CHANGELOG.md` を更新
- [x] テスト追加・更新
- [x] `uv run pytest -m "not com and not render"` を実行
- [x] `uv run task precommit-run` を実行

## Review

- Summary: `create_chart` と `apply_table_style` の同時指定を `service.py` で明示拒否するガードを追加。`runtime.py` に `contains_create_chart_op` を公開し、`test_service.py` に混在オペレーション拒否テストを追加。関連ドキュメント（`docs/mcp.md` / `README.md` / `README.ja.md` / `CHANGELOG.md` / `docs/README.en.md` / `docs/README.ja.md`）を更新。
- Risks: 実行時に mixed-op をエラー化したため、従来この組み合わせを1リクエストで投げていたクライアントは分割呼び出しが必要。
- Follow-ups: 必要なら `docs/mcp.md` に `create_chart` 専用のサンプル（line/column/pie）を追加する。
