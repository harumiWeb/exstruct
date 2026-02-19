# Feature Spec for AI Agent (Phase-by-Phase)

## 1. Feature
Issue: MCP tool for creating a new Excel workbook (`exstruct_make`)

## 2. Goal
`exstruct_make` により、新規Excelブック作成と `ops` 適用を1回の呼び出しで完結できるようにする。  
`exstruct_patch` は既存編集専用として責務を分離し、AIエージェントの誤用を減らす。

## 3. In Scope
- 新規MCPツール `exstruct_make` の追加
- 空Workbookを作成し、`ops` を適用して保存
- `out_path` 必須I/F
- `ops` は任意（空配列を許容）
- 拡張子 `.xlsx` / `.xlsm` / `.xls` を受け入れる
- `.xls` は COM 必須でサポート
- 初期シート名を `Sheet1` に正規化
- `dry_run` / `return_inverse_ops` / `preflight_formula_check` / `auto_formula` / `backend` を `exstruct_patch` 同等で提供
- 既存 `run_patch` ロジックを再利用して差分・エラー形式互換を維持
- server/tool/docs/tests を更新

## 4. Out of Scope
- テンプレート複製作成（`template_path` 等）
- `exstruct_patch` の挙動変更
- 既存 `PatchOp` スキーマの破壊的変更
- 新規op追加（`ops` 自体は既存定義をそのまま利用）

## 5. Public Interface Changes

### 5.1 New MCP tool: `exstruct_make`
- 入力（想定）:
  - `out_path: str`（必須）
  - `ops: list[PatchOp]`（任意）
  - `on_conflict: overwrite | skip | rename | None`
  - `auto_formula: bool = false`
  - `dry_run: bool = false`
  - `return_inverse_ops: bool = false`
  - `preflight_formula_check: bool = false`
  - `backend: auto | com | openpyxl = auto`
- 出力:
  - `PatchToolOutput` 互換

### 5.2 Internal API additions
- `MakeRequest` / `run_make` を追加
- `MakeToolInput` / `MakeToolOutput` / `run_make_tool` を追加
- `server.py` へ `exstruct_make` 登録を追加

## 6. Validation Rules
- `out_path` は必須
- 対応拡張子は `.xlsx` / `.xlsm` / `.xls` のみ
- `ops` は object配列を正準とし、JSON object文字列配列も受理
- `.xls` で `backend='openpyxl'` はエラー
- `.xls` で COM 非利用環境はエラー
- `.xls` で `dry_run` / `return_inverse_ops` / `preflight_formula_check` はエラー（COM経路制約）
- PathPolicy の root/deny_glob 制約を適用

## 7. Backend Behavior

### 7.1 `.xlsx` / `.xlsm`
- 空Workbookを生成（初期シート名は `Sheet1`）
- 生成したseedを入力として `run_patch` 再利用
- `backend=auto` は `run_patch` 既存選択ルールに従う

### 7.2 `.xls`
- COMでseed作成
- COM利用不可なら明示エラー
- `backend=auto`/`com` でCOM経路を使用
- openpyxl専用機能フラグは不許可

## 8. Compatibility
- `exstruct_patch` は既存編集専用のまま不変
- `PatchResult` / `PatchToolOutput` 構造を維持
- 既存テストの期待値を壊さない

## 9. Test Scenarios
- `ops=[]` で `.xlsx` を作成し `Sheet1` が存在する
- `ops` 付きで `add_sheet` / `set_value` が適用される
- `on_conflict` の `overwrite` / `skip` / `rename` が期待通り
- `out_path` が PathPolicy 外なら失敗
- `.xls` + COMなしで失敗
- `.xls` + `backend='openpyxl'` で失敗
- `.xls` + `dry_run` 等で失敗
- JSON文字列opsが正規化される
- server既定 `--on-conflict` が `exstruct_make` に伝播する
- 既存 `exstruct_patch` 回帰なし

## 10. Acceptance Criteria
- `exstruct_make` で新規作成と `ops` 適用が可能
- 異常系バリデーションが想定通り失敗
- `tests/mcp` の関連テストが通過
- `uv run task precommit-run` が通過

## 11. Assumptions / Defaults
- 新規作成起点は空Workbookのみ
- `out_path` 必須
- `ops` は任意（空許容）
- `.xls` は COM 必須で対応
- 初期シート名は常に `Sheet1` に正規化
- `exstruct_patch` と同等の拡張フラグを維持
