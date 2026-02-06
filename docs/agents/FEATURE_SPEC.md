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
    {"op": "set_value", "sheet": "フォーム", "cell": "B2", "value": "山田太郎"}
  ]
}
```

### set_formula
```json
{
  "xlsx_path": "book.xlsx",
  "ops": [
    {"op": "set_formula", "sheet": "計算", "cell": "C10", "formula": "=SUM(Sheet1!A:A)"}
  ]
}
```

### add_sheet
```json
{
  "xlsx_path": "book.xlsx",
  "ops": [
    {"op": "add_sheet", "sheet": "売上集計"},
    {"op": "set_value", "sheet": "売上集計", "cell": "A1", "value": "月"}
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
