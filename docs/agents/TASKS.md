# Task List

未完了 [ ], 完了 [x]

## MCPサーバー（MVP）

- [x] 仕様反映: `docs/agents/FEATURE_SPEC.md` を更新
- [x] 依存追加: `pyproject.toml` に `exstruct[mcp]` の extras を追加
- [x] エントリポイント: `exstruct-mcp = exstruct.mcp.server:main` を定義
- [x] MCP 基盤: `src/exstruct/mcp/server.py` を追加（stdio サーバー起動）
- [x] ツール定義: `src/exstruct/mcp/tools.py` に `exstruct_extract` を実装
- [x] パス制約: `src/exstruct/mcp/io.py` に allowlist / deny glob を実装
- [x] 抽出実行: `src/exstruct/mcp/extract_runner.py` に内部 API 優先の実行層を実装
- [x] 出力モデル: Pydantic で入出力モデルを定義（mypy strict / Ruff 遵守）
- [x] ログ: stderr / ファイル出力の設定を追加
- [x] ドキュメント: README または docs に起動例（`exstruct-mcp --root ...`）を追記

## MCPサーバー（実用化）

- [x] `exstruct_read_json_chunk` を追加（大容量 JSON 対応）
- [x] `exstruct_validate_input` を追加（事前検証）
- [x] `--on-conflict` の出力衝突ポリシー実装
- [x] Windows/非Windows の読み取り差分を明文化
- [x] 最低限のテスト追加（パス制約 / 入出力モデル / 例外）

## PR #47 レビュー対応

- [x] cells.py の列幅縮小ヒューリスティックを再検討（遅い行に境界があるケースで早期縮小しない方針に修正）
- [x] 上記修正に対応するテストを追加（遅い行・右端に表があるケースを openpyxl で検証）
- [x] Codecov 指摘の不足分を埋めるテスト追加（mcp: chunk_reader/extract_runner/server/tools/validate_input/io、core/cells）
- [x] CodeRabbit: Docstring coverage 80% を満たすよう不足分の docstring を追加
