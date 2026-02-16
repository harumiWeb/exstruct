# Feature Spec for AI Agent (Phase-by-Phase)

本ドキュメントは、ExStruct MCP に新規追加する編集ツール `exstruct_patch` の MVP 仕様を定義します。

---

## 1. 目的・スコープ

- **MVP対象**: セル値更新 / 数式更新 / 新規シート追加
- **非対象**: 図形・チャート編集、大規模レイアウト変更
- **再計算**: 行わない（Excel側の再計算に委ねる）

---

## 2. ツールI/F定義（OpenAPI / Pydantic想定）

### Tool名

- `exstruct_patch`

### Input

- `xlsx_path: str`
- `ops: list[PatchOp]`
- `out_dir: str | None`
- `out_name: str | None`
- `on_conflict: "overwrite" | "skip" | "rename" | None`
- `auto_formula: bool`（default: false）

### Output

- `out_path: str`
- `patch_diff: list[PatchDiffItem]`
- `warnings: list[str]`
- `error: PatchErrorDetail | null`

---

## 3. 操作種別（PatchOp）

### set_value

- `op: "set_value"`
- `sheet: str`
- `cell: str` （A1形式）
- `value: str | int | float | None`（nullでクリア）

### set_formula

- `op: "set_formula"`
- `sheet: str`
- `cell: str` （A1形式）
- `formula: str`（必ず `=` から開始）

### add_sheet

- `op: "add_sheet"`
- `sheet: str`（新規シート名）

---

## 4. バリデーション規則

- `xlsx_path`
  - 存在するファイルであること
  - 許可パス内であること（PathPolicy）
  - 拡張子は `.xlsx` / `.xlsm` / `.xls`（`.xls` は COM 利用可能環境のみ）
- `sheet`
  - 空文字禁止
  - `add_sheet` は既存名と重複禁止
  - `set_value` / `set_formula` は存在シートのみ
- `cell`
  - A1形式を必須（例: `B3`）
- `set_value`
  - `value` が `=` から始まる場合は通常は拒否
  - `auto_formula=true` のときは `set_formula` 相当として処理
- `set_formula`
  - `formula` は必ず `=` から開始

---

## 5. 実行セマンティクス

- opsは **順序通り** に評価・適用
- **全て成功した場合のみ保存**（原子性）
- 失敗時は **Excelを変更せずエラー** を返す
- `add_sheet` 実行後にのみ、そのシートへの `set_value` / `set_formula` を許可

---

## 6. 出力差分（patch_diff）

最低限の構造:

- `op: "set_value" | "set_formula" | "add_sheet"`
- `op_index: int`
- `sheet: str`
- `cell: str | null`
- `before: PatchValue | null`
- `after: PatchValue | null`
- `status: "applied" | "skipped"`

`PatchValue`:

- `kind: "value" | "formula" | "sheet"`
- `value: str | int | float | null`

`PatchErrorDetail`:

- `op_index: int`
- `op: "set_value" | "set_formula" | "add_sheet"`
- `sheet: str`
- `cell: str | null`
- `message: str`

---

## 7. バックエンド方針

- **Windows + COM利用可能**: COM優先
- **それ以外**: openpyxl
- openpyxl使用時は、図形/チャート等の保持制限を `warnings` で通知

---

## 8. 競合ポリシー

- `on_conflict` 未指定時は **サーバー起動時の `--on-conflict` 設定値** を既定値とする
- `skip` の場合:
  - `patch_diff` は空配列
  - `warnings` にスキップ理由を記載

出力名の既定:

- `out_name` 未指定時は `"{stem}_patched{suffix}"`

---

## 9. エラー処理

- パス違反 / セル不正 / 数式不正: `ValueError`
- シート不在 / 操作不正: `PatchErrorDetail` を `error` に格納
- 読み込み不能: `FileNotFoundError` / `OSError`
- バックエンド例外: `RuntimeError`

---

## 10. 例

### set_value

```json
{
  "xlsx_path": "book.xlsx",
  "ops": [
    {
      "op": "set_value",
      "sheet": "フォーム",
      "cell": "B2",
      "value": "山田太郎"
    }
  ]
}
```

### set_formula

```json
{
  "xlsx_path": "book.xlsx",
  "ops": [
    {
      "op": "set_formula",
      "sheet": "計算",
      "cell": "C10",
      "formula": "=SUM(Sheet1!A:A)"
    }
  ]
}
```

### add_sheet

```json
{
  "xlsx_path": "book.xlsx",
  "ops": [
    { "op": "add_sheet", "sheet": "売上集計" },
    { "op": "set_value", "sheet": "売上集計", "cell": "A1", "value": "月" }
  ]
}
```

---

## 11. レビュー対応仕様（MVP修正）

### 11.1 `on_conflict` 既定値の統一

- サーバーCLI引数 `--on-conflict` の既定値・指定値は、`exstruct_extract` と `exstruct_patch` の双方に同一に適用する
- `exstruct_patch` で固定値 `rename` を持たない
- ツール呼び出しで `on_conflict` が未指定の場合:
  - サーバー起動時の `--on-conflict` を適用
- ツール呼び出しで `on_conflict` が指定された場合:
  - ツール引数を最優先する

### 11.2 出力ディレクトリ生成の明確化

- `exstruct_patch` は保存前に `out_path.parent.mkdir(parents=True, exist_ok=True)` 相当を実行する
- `out_dir` が存在しない場合でも保存可能にする

### 11.3 `.xls` サポート条件の明確化

- `.xls` は **COM利用可能環境のみ編集対象** とする
- COM利用不可時の `.xls` は `ValueError` を返す（メッセージで「COM必須」を明示）
- バリデーション規則の拡張子記述を次に更新する:
  - `.xlsx` / `.xlsm`: 常時対象
  - `.xls`: 条件付き対象（Windows + COM）

### 11.4 後方互換性

- 既存の `PatchOp` / `PatchResult` スキーマは変更しない
- 変更はデフォルト解決順序・保存前処理・エラーメッセージの明確化に限定する

### 11.5 テスト要件（レビュー対応分）

- `exstruct_patch` で `on_conflict` 未指定時にサーバー既定値が反映されること
- `exstruct_patch` で `on_conflict` 指定時にツール引数が優先されること
- 未作成 `out_dir` 指定で保存成功すること
- `.xls` + COM不可時に期待どおり `ValueError` となること（メッセージ確認を含む）

---

## 12. 次フェーズ仕様（便利機能 1〜5）

### 12.1 対象機能

- `dry_run`（非破壊プレビュー）
- `undo` 用逆パッチ生成
- 範囲操作（値一括設定 / 数式フィル）
- 条件付き更新（CAS）
- 数式ヘルスチェック

### 12.2 ツールI/F拡張

`exstruct_patch` の Input に以下を追加:

- `dry_run: bool = false`
- `return_inverse_ops: bool = false`
- `preflight_formula_check: bool = false`

`exstruct_patch` の Output に以下を追加:

- `inverse_ops: list[PatchOp]`（default: `[]`）
- `formula_issues: list[FormulaIssue]`（default: `[]`）

`FormulaIssue`:

- `sheet: str`
- `cell: str`
- `level: "warning" | "error"`
- `code: "invalid_token" | "ref_error" | "name_error" | "div0_error" | "value_error" | "na_error" | "circular_ref_suspected"`
- `message: str`

### 12.3 操作種別拡張（PatchOp）

#### set_range_values

- `op: "set_range_values"`
- `sheet: str`
- `range: str`（A1範囲形式、例: `B2:D4`）
- `values: list[list[str | int | float | None]]`
- `shape(values)` と `range` サイズが一致しない場合は `ValueError`

#### fill_formula

- `op: "fill_formula"`
- `sheet: str`
- `range: str`（縦または横の連続範囲）
- `base_cell: str`
- `formula: str`（`=` 始まり）
- `formula` を `base_cell` に置いたときの相対参照を、`range` 内に展開する

#### set_value_if

- `op: "set_value_if"`
- `sheet: str`
- `cell: str`
- `expected: str | int | float | None`
- `value: str | int | float | None`
- 現在値が `expected` と一致した場合のみ更新
- 不一致時は `PatchDiffItem.status="skipped"` とし、`warnings` に理由を追加

#### set_formula_if

- `op: "set_formula_if"`
- `sheet: str`
- `cell: str`
- `expected: str | int | float | None`
- `formula: str`（`=` 始まり）
- 条件判定は `set_value_if` と同様

### 12.4 実行セマンティクス（追加）

- `dry_run=true` の場合:
  - 保存しない
  - `patch_diff` / `warnings` / `formula_issues` / `inverse_ops` のみ返す
  - `out_path` は保存予定パスを返す
- `return_inverse_ops=true` の場合:
  - `applied` な変更のみ逆操作を生成する
  - 逆操作順は「適用順の逆順」とする
- `preflight_formula_check=true` の場合:
  - 適用後ワークブックを走査し `formula_issues` を返す
  - `level=error` がある場合の扱い:
    - `dry_run=false`: 変更は保存せず `error` を返す
    - `dry_run=true`: 保存なしのため `error` は返さず、問題一覧のみ返す

### 12.5 バリデーション規則（追加）

- `range` は A1 範囲形式（例: `A1:C3`）必須
- `fill_formula.range` は単一行または単一列のみ（MVP制約）
- 条件付き更新の `expected` 比較は「内部正規化後の値一致」で判定
- `set_value_if` で `value` が `=` 始まりかつ `auto_formula=false` の場合は既存 `set_value` と同様に拒否

### 12.6 後方互換性

- 既存 `set_value` / `set_formula` / `add_sheet` の仕様は変更しない
- 新規フィールドは optional 追加のみで、既存クライアントを破壊しない

### 12.7 受け入れ基準

- `dry_run=true` でファイル更新時刻が変化しない
- `return_inverse_ops=true` で、逆パッチ適用により元状態へ戻せる
- `set_range_values` がサイズ不一致を検知できる
- `set_value_if` / `set_formula_if` が不一致時に `skipped` を返せる
- `preflight_formula_check=true` で `#REF!` 等を検出して `formula_issues` に反映できる

---

## 13. ABC列名キー出力オプション（`alpha_col`）

### 13.1 背景・目的

AIエージェントがExStruct MCPで抽出した結果をもとに `exstruct_patch` でセル編集を行う際、
`CellRow.c` のキーが 0-based 数値文字列（`"0"`, `"12"`, `"18"` 等）のため、
Excel上の列名（A, M, S 等）との対応が直感的に取れず、**誤ったセル座標への書き込みが頻発する**。

本機能は、抽出結果の `CellRow.c` キーを Excel 列名形式（A, B, ..., Z, AA, AB, ...）で
出力するオプションを提供し、AIエージェントの座標理解精度を大幅に改善する。

### 13.2 スコープ

- **対象**: `CellRow.c` のキー、`CellRow.links` のキー
- **非対象**: `MergedCells.items` / `PrintArea` / `formulas_map` / `colors_map` の列インデックス（整数のまま）
- **変換方式**: 0-based 数値 → Excel列名（`0` → `A`, `1` → `B`, `25` → `Z`, `26` → `AA`）

### 13.3 I/F

#### 本体 API（`ExStructEngine` / `extract()` / `process_excel()`）

- `StructOptions` に `alpha_col: bool = False` を追加
- `True` の場合、抽出結果の `CellRow.c` / `CellRow.links` のキーを ABC 形式に変換する

#### CLI

- `--alpha-col` フラグを追加（`action="store_true"`）
- `process_excel()` に `alpha_col` パラメータを追加

#### MCP ツール（`exstruct_extract`）

- `ExtractOptions` に `alpha_col: bool = False` を追加
- `ExtractToolInput.options.alpha_col` をエンジンに連携

### 13.4 変換ユーティリティ

- `exstruct.models.col_index_to_alpha(index: int) -> str` を新設
  - 0-based 整数 → Excel 列名文字列
  - 例: `0` → `"A"`, `25` → `"Z"`, `26` → `"AA"`, `701` → `"ZZ"`
- `exstruct.models.convert_row_keys_to_alpha(row: CellRow) -> CellRow` を新設
  - `CellRow` の `c` / `links` キーを一括変換した新しい `CellRow` を返す
- `exstruct.models.convert_sheet_keys_to_alpha(sheet: SheetData) -> SheetData` を新設
  - `SheetData.rows` 内の全 `CellRow` を変換した新しい `SheetData` を返す

### 13.5 変換タイミング

- `extract_workbook()` でデータ抽出が完了した **後**、シリアライズの **前** に変換する
- パイプライン上の位置: `ExStructEngine.extract()` の返却直前

### 13.6 出力例

#### `alpha_col=False`（デフォルト: 従来通り）

```json
{ "r": 6, "c": { "0": "氏　名", "12": "生年月日", "18": "年" } }
```

#### `alpha_col=True`

```json
{ "r": 6, "c": { "A": "氏　名", "M": "生年月日", "S": "年" } }
```

### 13.7 後方互換性

- デフォルト `False` のため、既存の出力は一切変更されない
- `alpha_col=True` は opt-in のみ
- `_safe_col_index()` 等の下流処理は数値キーを前提としているが、変換はシリアライズ直前のため影響なし

### 13.8 受け入れ基準

- `col_index_to_alpha(0)` → `"A"`, `col_index_to_alpha(25)` → `"Z"`, `col_index_to_alpha(26)` → `"AA"`
- `alpha_col=True` で抽出した `CellRow.c` のキーがすべて A-Z 形式
- `alpha_col=True` で `CellRow.links` のキーも同様に変換される
- `alpha_col=False`（デフォルト）で従来通り数値キー
- CLI `--alpha-col` フラグが正しく `process_excel()` に伝播する
- MCP `exstruct_extract` の `options.alpha_col` がエンジンに伝播する

### 13.9 `merged_ranges` 出力（`alpha_col` 時）

#### 背景

`alpha_col=True` 時に `CellRow.c` / `links` は ABC 列名になるため、`merged_cells.items`
（`(r1, c1, r2, c2, v)` の数値列インデックス）との参照体系が混在する。
AIエージェントの編集座標推論の一貫性を高めるため、A1形式の `merged_ranges` を別フィールドで出力する。

#### 仕様

- `SheetData` に `merged_ranges: list[str]` を追加する（例: `["A1:C3", "F10:G12"]`）
- `alpha_col=True` のときのみ `merged_ranges` を生成する
- `alpha_col=True` では `merged_cells` は出力しない（可読性優先）
- `alpha_col=False` のときは従来どおり `merged_cells` を出力し、`merged_ranges` は空とする

#### 受け入れ基準（追加）

- `alpha_col=True` かつ `merged_cells` が存在するシートで `merged_ranges` が出力される
- `merged_ranges` の各要素は `"A1:C3"` 形式で、`items` と同じ範囲を表す
- `alpha_col=True` では `merged_cells` は出力されない
- `alpha_col=False` では `merged_ranges` は空（実質非出力）となる

---

## 14. `exstruct_patch.ops` 入力スキーマ整合（AIエージェント運用改善）

### 14.1 背景

一部のMCPクライアントは、`ops` を `PatchOp` オブジェクト配列ではなく、
JSON文字列配列として送信する場合がある。
この差異により、AIエージェントがツール定義どおりに呼び出しても
入力バリデーションで失敗するケースがある。

### 14.2 目的

- `exstruct_patch` の受け口を明確化し、実装とツール入力スキーマの体験差を解消する
- AIエージェントがクライアント差異を意識せず patch を実行できるようにする

### 14.3 入力仕様

- 正式入力: `ops: list[object]`（各要素は `PatchOp` 相当のJSONオブジェクト）
- 互換入力: `ops: list[str]`（各要素は `PatchOp` オブジェクトを表すJSON文字列）
- サーバー内部で次を行う:
  - object はそのまま採用
  - string は `json.loads()` で object に変換
  - 変換後に `PatchToolInput`（`ops: list[PatchOp]`）で厳密検証

### 14.4 エラー仕様

- JSON文字列のパース失敗時:
  - `ValueError`
  - メッセージに失敗した `ops[index]` を含める
  - 「object 形式の推奨入力例」を併記する
- JSONとして有効でもobjectでない場合（配列/数値など）:
  - `ValueError`
  - 「JSON object 必須」を明記する

### 14.5 後方互換性

- 既存の object 配列クライアントは変更不要
- string 配列クライアントはサーバー側で吸収し、破壊的変更を回避する
- `PatchOp` 自体のスキーマと patch 実行ロジックは変更しない

### 14.6 受け入れ基準

- object 配列入力で従来どおり patch が実行できる
- JSON文字列配列入力でも patch が実行できる
- 不正JSON文字列で `ops[index]` を含む明確なエラーが返る
- object 以外のJSON型入力で「object必須」エラーが返る

---

## 15. MCP モード/チャンク指定ガイド改善

### 15.1 背景

MCP運用時、次の迷いが発生しやすい。

- `exstruct_extract.mode` に無効値を入れて失敗する
- `exstruct_read_json_chunk` で `sheet` 未指定のまま再試行して失敗する
- `max_bytes` の調整方針が不明でリトライ回数が増える
- `alpha_col=true` 出力時に `filter.cols` が使いにくい

### 15.2 目的

- モード指定とチャンク指定の「最短成功パターン」を文書化する
- エラー時に次の行動が分かるメッセージへ改善する
- `alpha_col=true` でも列フィルタを直感的に使えるようにする

### 15.3 スコープ

- `docs/mcp.md` の実運用ガイド追加
- `src/exstruct/mcp/server.py` のツールDocstring詳細化
- `src/exstruct/mcp/chunk_reader.py` のエラーメッセージ改善
- `src/exstruct/mcp/chunk_reader.py` の `filter.cols` を英字列キー対応
- 関連テスト追加

### 15.4 仕様

#### 15.4.1 ドキュメント強化（`docs/mcp.md`）

- `validate_input -> extract(standard) -> read_json_chunk` の推奨3ステップを追加
- `mode` の許容値（`light`/`standard`/`verbose`）と用途を表形式で明記
- `read_json_chunk` の主要引数（`sheet`, `max_bytes`, `filter`, `cursor`）を具体化
- エラー/警告ごとの再試行手順を表形式で明記

#### 15.4.2 ツール説明の詳細化（`server.py`）

- `exstruct_extract` Docstringに `mode` の意味と許容値を明記
- `exstruct_read_json_chunk` Docstringに `filter.rows/cols` の
  1-based inclusive 仕様と `max_bytes` 調整指針を明記

#### 15.4.3 エラーメッセージ改善（`chunk_reader.py`）

- 大きすぎる生JSONに対するエラー文言を、`sheet`/`filter` 指定例付きへ変更
- 複数シート時の `sheet` 必須エラーに、利用可能シート名と再実行ヒントを追加

#### 15.4.4 `filter.cols` の英字列キー対応（`chunk_reader.py`）

- 列キーが数値文字列（`"0"`）に加え英字列名（`"A"`, `"AA"`）でも判定可能にする
- `filter.cols` は従来どおり 1-based inclusive の列番号指定を維持する
- 英字列名は大文字小文字を区別しない

### 15.5 後方互換性

- 既存の `filter.cols` 数値キー挙動は維持する
- `read_json_chunk` の入出力スキーマは変更しない
- 変更はガイド追加・メッセージ改善・列キー判定拡張に限定する

### 15.6 受け入れ基準

- `docs/mcp.md` だけで、初回利用者が mode/chunk 再試行を完了できる
- `Output is too large` エラーに次アクションが含まれる
- `Sheet is required when multiple sheets exist` エラーにシート候補が含まれる
- `filter.cols` が数値キー・英字キー双方で機能する
- 既存テストと追加テストが通過する
