# Feature Spec

## Feature Name

Issue 72: MCP `exstruct_capture_sheet_images` (範囲指定付き画像エクスポート)

## Goal

AIエージェントがExcelシートを視覚検証できるように、MCPツールからPNG画像を安全に出力できるようにする。

- 範囲指定（`A1:B2` / `Sheet1!A1:B2` / `'Sheet 1'!A1:B2`）を受け付ける
- COM限定で動作させる
- `copyPicture` は使わず、PDF変換 + `pypdfium2` で画像化する
- `out_dir` 未指定時は root 直下に一意ディレクトリを作る
- `out_dir` 指定有無に関係なく、実行後に出力先ディレクトリパスを返す

## Scope

### In Scope

- MCP新規ツール `exstruct_capture_sheet_images` 追加
- `sheet` / `range` 入力に基づく対象シート・対象範囲の画像出力
- root配下の一意出力ディレクトリ解決
- COM可用性チェック
- 関連テスト追加・更新
- docs更新

### Out of Scope

- `copyPicture` ベース実装
- openpyxlのみでの画像出力
- 既存CLI引数の拡張
- PDF以外の中間形式追加

## Public API / Interface Changes

### MCP Tool (new)

- Tool name: `exstruct_capture_sheet_images`
- Input:
  - `xlsx_path: str` (required)
  - `out_dir: str | None = None`
  - `dpi: int = 144` (`>= 1`)
  - `sheet: str | None = None`
  - `range: str | None = None`
- Output:
  - `out_dir: str` (resolved output directory, always returned)
  - `image_paths: list[str]`
  - `warnings: list[str]`

### Python Internal API Changes

- `src/exstruct/render/__init__.py`
  - `export_sheet_images(...)` に以下を追加:
    - `sheet: str | None = None`
    - `a1_range: str | None = None`

### Internal Models (new)

- `CaptureSheetImagesRequest` (Pydantic BaseModel)
  - `xlsx_path: Path`
  - `out_dir: Path | None = None`
  - `dpi: int = Field(default=144, ge=1)`
  - `sheet: str | None = None`
  - `range: str | None = None`
- `CaptureSheetImagesResult` (Pydantic BaseModel)
  - `out_dir: str`
  - `image_paths: list[str]`
  - `warnings: list[str] = Field(default_factory=list)`

## Behavior Spec

### 1. `sheet` / `range` のルール

- `range is None`:
  - `sheet is None` の場合は全シート出力
  - `sheet` 指定時は指定シートのみ出力
- `range is not None`:
  - `sheet` 必須（未指定はエラー）
  - `range` は以下を許可:
    - `A1:B2`
    - `Sheet1!A1:B2`
    - `'Sheet 1'!A1:B2`
  - `sheet` と `range` のシート名が不一致ならエラー

### 2. 出力先ディレクトリ

- `out_dir` 指定あり:
  - 指定先を利用（`PathPolicy` で許可範囲検証）
- `out_dir` 未指定:
  - MCP `--root` 直下に `<workbook_stem>_images` を候補作成
  - 競合時は `<workbook_stem>_images_1`, `_2`, ... で一意化
- 実行結果には常に `out_dir`（解決済み実パス）を含める

### 3. バックエンド/実装制約

- COM限定機能（Excel COM利用不可ならエラー）
- 画像生成パイプラインは現行維持:
  - `ExportAsFixedFormat(PDF)` -> `pypdfium2` -> PNG
- `copyPicture` は採用しない

## Error Handling

- `dpi < 1` -> ValidationError
- `range` 形式不正 -> ValueError
- `range` 指定時の `sheet` 未指定 -> ValueError
- `sheet` と `range` シート修飾不一致 -> ValueError
- COM不可 -> ValueError (COM必須メッセージ)
- 出力先が `PathPolicy` 範囲外 -> ValueError
- PDF/画像変換失敗 -> RenderError

## Files to Change

- `src/exstruct/mcp/server.py`
- `src/exstruct/mcp/tools.py`
- `src/exstruct/mcp/render_runner.py` (new)
- `src/exstruct/mcp/shared/output_path.py`
- `src/exstruct/mcp/shared/a1.py`
- `src/exstruct/mcp/__init__.py`
- `src/exstruct/render/__init__.py`
- `tests/mcp/test_tool_models.py`
- `tests/mcp/test_tools_handlers.py`
- `tests/mcp/test_server.py`
- `tests/mcp/shared/test_output_path.py`
- `tests/render/test_render_init.py`
- `docs/mcp.md`
- `README.md`
- `README.ja.md`

## Test Scenarios

### Model/Validation

- `dpi=0` を拒否
- `range` 指定時の `sheet` 未指定を拒否
- `range` のシート修飾付き形式を受理・正規化
- `sheet` / `range` 不一致を拒否

### Tool Handler/Server

- `exstruct_capture_sheet_images` が登録される
- 入力が `run_capture_sheet_images_tool` に正しく渡る
- 出力 `out_dir` / `image_paths` が返る

### Output Path

- 未指定 `out_dir` で `<stem>_images` 作成
- 衝突時に連番リネーム
- 指定 `out_dir` はそのまま使う

### Render

- `sheet + range` 指定で対象のみ出力
- `sheet` のみ指定で単一シート出力
- `range` なしで既存動作を維持
- 不正範囲入力時に失敗する

## Acceptance Criteria

- MCPツール `exstruct_capture_sheet_images` が利用できる
- 範囲指定仕様（A1/B2、シート修飾、不一致エラー）が満たされる
- `out_dir` 未指定時に root直下へ一意ディレクトリ作成される
- レスポンスに常に `out_dir` が含まれる
- COM限定・PDF->PNG経路が守られる
- 追加テストと既存回帰テストが通る

## Verification Commands

- `uv run pytest tests/mcp/test_tool_models.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/shared/test_output_path.py tests/render/test_render_init.py`
- `uv run task precommit-run`
