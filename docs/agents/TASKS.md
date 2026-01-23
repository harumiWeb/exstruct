# Task List

未完了 [ ], 完了 [x]

## MCPサーバー（MVP）

- [x] 仕様反映: `docs/agents/FEATURE_SPEC.md` を更新
- [ ] 依存追加: `pyproject.toml` に `exstruct[mcp]` の extras を追加
- [ ] エントリポイント: `exstruct-mcp = exstruct.mcp.server:main` を定義
- [ ] MCP 基盤: `src/exstruct/mcp/server.py` を追加（stdio サーバー起動）
- [ ] ツール定義: `src/exstruct/mcp/tools.py` に `exstruct.extract` を実装
- [ ] パス制約: `src/exstruct/mcp/io.py` に allowlist / deny glob を実装
- [ ] 抽出実行: `src/exstruct/mcp/extract_runner.py` に内部 API 優先の実行層を実装
- [ ] 出力モデル: Pydantic で入出力モデルを定義（mypy strict / Ruff 遵守）
- [ ] ログ: stderr / ファイル出力の設定を追加
- [ ] ドキュメント: README または docs に起動例（`exstruct-mcp --root ...`）を追記

## MCPサーバー（実用化）

- [ ] `exstruct.read_json_chunk` を追加（大容量 JSON 対応）
- [ ] `exstruct.validate_input` を追加（事前検証）
- [ ] `--on-conflict` の出力衝突ポリシー実装
- [ ] Windows/非Windows の読み取り差分を明文化
- [ ] 最低限のテスト追加（パス制約 / 入出力モデル / 例外）
