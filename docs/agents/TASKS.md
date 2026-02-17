# Task List

未完了 [ ], 完了 [x]

- [x] `FEATURE_SPEC.md` に exstruct_patch の詳細仕様（I/F、検証、差分、エラー）を記載
- [x] `mcp/patch_runner.py` を新設し、`PatchRequest/Result` を定義
- [x] `mcp/tools.py` に `PatchToolInput/Output` と `run_patch_tool` を追加
- [x] `mcp/server.py` に `exstruct_patch` 登録を追加
- [x] パス制御・競合ポリシー・保存ロジックを実装
- [x] openpyxlベースの適用処理を実装
- [x] COM利用時の分岐（可能なら）とwarning出力
- [x] ユニットテスト（非COM）を追加

## レビュー対応タスク（MVP修正）

- [x] `mcp/server.py`: `exstruct_patch` の `on_conflict` 解決で固定 `rename` を廃止し、サーバー既定値（`default_on_conflict`）を利用する
- [x] `mcp/patch_runner.py`: 保存前に出力先ディレクトリ作成処理を追加する（未作成 `out_dir` 対応）
- [x] `mcp/patch_runner.py`: `.xls` + COM不可時のエラーメッセージを「COM必須」が明確な文言に調整する
- [x] `docs/agents/FEATURE_SPEC.md`: 拡張子要件を「`.xls` はCOM利用可能環境のみ」に更新する（本文整合）
- [x] `tests/mcp/test_server.py`: patch ツールの `on_conflict` 既定値伝播を検証するテストを追加/更新する
- [x] `tests/mcp/test_patch_runner.py`: 未作成 `out_dir` 保存成功ケースを追加する
- [x] `tests/mcp/test_patch_runner.py`: `.xls` + COM不可エラーの期待値（型・文言）を追加する
- [x] `ruff check .` / `mypy src/exstruct --strict` / 関連pytest を実行し、結果を記録する

## 次フェーズタスク（便利機能 1〜5）

- [x] `mcp/patch_runner.py`: `PatchRequest` に `dry_run` / `return_inverse_ops` / `preflight_formula_check` を追加する
- [x] `mcp/patch_runner.py`: `PatchResult` に `inverse_ops` / `formula_issues` を追加する
- [x] `mcp/patch_runner.py`: `PatchOp` に `set_range_values` / `fill_formula` / `set_value_if` / `set_formula_if` を追加する
- [x] `mcp/patch_runner.py`: A1範囲パース・範囲サイズ検証・値行列検証ユーティリティを実装する
- [x] `mcp/patch_runner.py`: `set_range_values` の openpyxl 実装を追加する
- [x] `mcp/patch_runner.py`: `fill_formula` の相対参照展開ロジックを実装する（単一行/列制約）
- [x] `mcp/patch_runner.py`: `set_value_if` / `set_formula_if` の比較・`skipped` 差分出力を実装する
- [x] `mcp/patch_runner.py`: 逆パッチ生成ロジック（`applied` のみ、逆順）を実装する
- [x] `mcp/patch_runner.py`: `dry_run=true` で非保存実行に分岐する
- [x] `mcp/patch_runner.py`: 数式ヘルスチェック（`#REF!` 等）を実装し、`formula_issues` に反映する
- [x] `mcp/patch_runner.py`: `preflight_formula_check=true` かつ `level=error` の保存抑止ルールを実装する
- [x] `mcp/tools.py`: `PatchToolInput/Output` の新規フィールドをツールI/Fへ公開する
- [x] `mcp/server.py`: `exstruct_patch` ツール引数に新規3フラグを追加し、ハンドラへ連携する
- [x] `docs/agents/FEATURE_SPEC.md`: 実装差分（命名、エラーコード、制約）を最終反映する
- [x] `tests/mcp/test_patch_runner.py`: `dry_run` 非保存・逆パッチ・範囲操作・条件付き更新・数式検査の単体テストを追加する
- [x] `tests/mcp/test_tools_handlers.py`: 入出力モデル拡張の受け渡しテストを追加する
- [x] `tests/mcp/test_server.py`: 新規ツール引数の受け渡しテストを追加する
- [x] `ruff check .` / `mypy src/exstruct --strict` / 関連pytest を実行し、結果を記録する

## ABC列名キー出力オプション（`alpha_col`）

- [x] `models/__init__.py`: `col_index_to_alpha()` 変換ユーティリティを追加する
- [x] `models/__init__.py`: `convert_row_keys_to_alpha()` / `convert_sheet_keys_to_alpha()` を追加する
- [x] `engine.py`: `StructOptions` に `alpha_col: bool = False` を追加する
- [x] `engine.py`: `ExStructEngine.extract()` で `alpha_col=True` 時に変換を適用する
- [x] `__init__.py`: `extract()` / `process_excel()` に `alpha_col` パラメータを追加する
- [x] `cli/main.py`: `--alpha-col` CLI フラグを追加する
- [x] `mcp/extract_runner.py`: `ExtractOptions` に `alpha_col` を追加し、`process_excel()` へ連携する
- [x] `mcp/server.py`: `exstruct_extract` ツールに `alpha_col` パラメータを追加する
- [x] `tests/models/`: `col_index_to_alpha()` の単体テストを追加する
- [x] `tests/models/`: `convert_row_keys_to_alpha()` / `convert_sheet_keys_to_alpha()` のテストを追加する
- [x] `tests/engine/`: `alpha_col=True` でのエンジン抽出テストを追加する
- [x] `tests/cli/`: `--alpha-col` CLI フラグのテストを追加する
- [x] `tests/mcp/`: MCP extract の `alpha_col` 連携テストを追加する
- [x] `ruff check .` / `mypy src/exstruct --strict` / 関連pytest を実行し、結果を記録する

## `merged_ranges` 出力（`alpha_col`時）

- [x] `docs/agents/FEATURE_SPEC.md`: `merged_ranges` 仕様に更新する
- [x] `models/__init__.py`: `SheetData` に `merged_ranges` フィールドを追加する
- [x] `models/__init__.py`: `alpha_col=True` 変換時に `merged_cells.items` から `merged_ranges` を生成する
- [x] `models/__init__.py`: `alpha_col=True` 変換時に `merged_cells` を非表示化する
- [x] `tests/models/test_alpha_col.py`: `merged_ranges` 生成と非生成（alpha_col=False相当）を検証する
- [x] `tests/engine/test_engine_alpha_col.py`: エンジン経由で `merged_ranges` が出力されることを検証する
- [x] `uv run task precommit-run` を実行し、静的解析・型検査の通過を確認する

## `exstruct_patch.ops` 入力スキーマ整合（AIエージェント運用改善）

- [x] `docs/agents/FEATURE_SPEC.md`: `ops` の正式入力（object配列）と互換入力（JSON文字列配列）を定義する
- [x] `mcp/server.py`: `ops` 正規化ヘルパー（object/string両対応）を追加する
- [x] `mcp/server.py`: `_patch_tool` で正規化済み `ops` を `PatchToolInput` に渡す
- [x] `mcp/server.py`: 不正JSON時に `ops[index]` と入力例を含むエラー文言を返す
- [x] `tests/mcp/test_server.py`: `ops` 正規化ヘルパーの単体テスト（成功/失敗）を追加する
- [x] `tests/mcp/test_server.py`: `exstruct_patch` が JSON文字列配列 `ops` を受理できることを検証する
- [x] `tests/mcp/test_server.py`: 不正JSON文字列 `ops` が明確な `ValueError` になることを検証する
- [x] `ruff check src/exstruct/mcp/server.py tests/mcp/test_server.py` を実行し、静的解析を確認する
- [x] `pytest tests/mcp/test_server.py` を実行し、回帰がないことを確認する

## MCP モード/チャンク指定ガイド改善

- [x] `docs/agents/FEATURE_SPEC.md`: mode/chunk ガイド改善仕様（目的/スコープ/受け入れ基準）を定義する
- [x] `docs/mcp.md`: 推奨3ステップ（validate/extract/chunk）と mode 比較表を追加する
- [x] `docs/mcp.md`: chunk パラメータ説明、エラー別リトライ表、cursor利用手順を追加する
- [x] `src/exstruct/mcp/server.py`: `exstruct_extract` Docstring に mode の許容値と意図を追記する
- [x] `src/exstruct/mcp/server.py`: `exstruct_read_json_chunk` Docstring に `filter.rows/cols` と `max_bytes` 指針を追記する
- [x] `src/exstruct/mcp/chunk_reader.py`: 大容量時エラー文言に再実行ヒントを追加する
- [x] `src/exstruct/mcp/chunk_reader.py`: 複数シート時エラー文言に利用可能シート名を追加する
- [x] `src/exstruct/mcp/chunk_reader.py`: `filter.cols` を英字列キー（`A`, `AA`）対応に拡張する
- [x] `tests/mcp/test_chunk_reader.py`: 英字列キー対応と新エラーメッセージのテストを追加する
- [x] `uv run task precommit-run` を実行し、静的解析・型検査・テスト通過を確認する

## 読み取りツール拡張（`read_range` / `read_cells` / `read_formulas`）

- [x] `docs/agents/FEATURE_SPEC.md`: 新規3ツールのI/F・バリデーション・受け入れ基準を定義する
- [x] `src/exstruct/mcp/sheet_reader.py`: A1セル/範囲パーサーと共通読み取りロジックを実装する
- [x] `src/exstruct/mcp/sheet_reader.py`: `read_range`（A1範囲指定）を実装する
- [x] `src/exstruct/mcp/sheet_reader.py`: `read_cells`（複数セル指定）を実装する
- [x] `src/exstruct/mcp/sheet_reader.py`: `read_formulas`（数式一覧＋任意で値）を実装する
- [x] `src/exstruct/mcp/tools.py`: `ReadRangeToolInput/Output`、`ReadCellsToolInput/Output`、`ReadFormulasToolInput/Output` を追加する
- [x] `src/exstruct/mcp/tools.py`: `run_read_range_tool`、`run_read_cells_tool`、`run_read_formulas_tool` を追加する
- [x] `src/exstruct/mcp/server.py`: `exstruct_read_range`、`exstruct_read_cells`、`exstruct_read_formulas` を登録する
- [x] `src/exstruct/mcp/server.py`: Docstringに利用例（`A1:G10`、`[\"J98\",\"J124\"]`）を追加する
- [x] `tests/mcp/test_sheet_reader.py`: A1パース、範囲上限、alpha_col互換、エラー系テストを追加する
- [x] `tests/mcp/test_tools_handlers.py`: 新規3ツールの入出力ハンドリングテストを追加する
- [x] `tests/mcp/test_server.py`: 新規3ツールの登録・引数連携テストを追加する
- [x] `docs/mcp.md`: 新規3ツールの使用手順と推奨フロー（extract→read_*）を追記する
- [x] `uv run task precommit-run` を実行し、静的解析・型検査・テスト通過を確認する
