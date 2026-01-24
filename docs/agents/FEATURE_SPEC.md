# Feature Spec for AI Agent (Phase-by-Phase)

本ドキュメントは AI エージェント向けに、段階的に実装を進めるための仕様メモです。

---

## MCPサーバー機能追加

### 目的

- MCP クライアント（Codex / Claude / VS Code Copilot / Gemini CLI 等）から ExStruct を「ツール」として安全に呼び出せるようにする
- 推論はエージェント側で行い、MCP は制御面（実行・結果参照）に徹する

### スコープ（MVP）

- stdio トランスポートの MCP サーバー
- ツール: `exstruct_extract`
- 抽出結果は **必ずファイル出力**（MCP 応答はパス + 軽いメタ情報）
- 安全なパス制約（allowlist / deny glob）

### 前提・制約

- 1MB 程度の Excel を想定
- 処理時間は長くなっても高品質重視
- Windows 以外は COM なしの簡易読み取り（ライブラリのスタンスに準拠）

### 出力 JSON の仕様

- `mode` で出力粒度を選択: `light` / `standard` / `verbose`
- 互換方針: 追加は OK、破壊的変更は NG

#### `light`

- 軽量メタデータ中心（シート名、件数、主要範囲など）
- 大きなセル本文や詳細構造は含めない

#### `standard`

- 通常運用向けの基本情報
- セル情報は要約・圧縮前提

#### `verbose`

- 詳細な構造情報を含む
- 大容量になりやすいため、ファイル出力＋チャンク取得前提

### MCP ツール仕様（案）

#### `exstruct_extract`

- 入力: `xlsx_path`, `mode`, `format`, `out_dir?`, `out_name?`, `options?`
- 出力: `out_path`, `workbook_meta`, `warnings`, `engine`
- 実装: 内部 API を優先、フォールバックで CLI サブプロセス

#### `exstruct_read_json_chunk`（実用化フェーズ）

- 入力: `out_path`, `sheet?`, `max_bytes?`, `filter?`, `cursor?`
- 出力: `chunk`, `next_cursor?`
- 方針: 返却サイズを抑制し、段階的に取得できること

#### `exstruct_validate_input`（実用化フェーズ）

- 入力: `xlsx_path`
- 出力: `is_readable`, `warnings`, `errors`

### サーバー設計

- stdio 優先
- ログは stderr / ファイル（stdio を汚さない）
- `--root` によりアクセス範囲を固定
- `--deny-glob` により防御的に除外
- `--on-conflict` で出力衝突方針を指定（overwrite / skip / rename）

### ディレクトリ構成（案）

```
src/exstruct/
  mcp/
    __init__.py
    server.py          # MCP server entrypoint (stdio)
    tools.py           # tool definitions + handlers
    io.py              # path validation, safe read/write
    extract_runner.py  # internal API call or subprocess fallback
    chunk_reader.py    # JSON partial read / pointer / sheet filters
```

---

## 今後のオプション検討メモ

- 表検知スコアリングの閾値を CLI/環境変数で調整可能にする
- 出力モード（light/standard/verbose）に応じてテーブル候補数を制限するオプション

---

## 実装方針

- 小さなステップごとにテスト追加、または既存フィクスチャで手動確認
- 短い関数・責務分割でスコアリング調整をしやすくする
- 外部公開前なので、破壊的変更はコメントや仕様に明示して段階的に移行する
