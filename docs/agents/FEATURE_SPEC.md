# Feature Spec for AI Agent (Phase-by-Phase)

## 1. Feature
Issue #61: MCPサーバーにExcelデザイン編集機能を追加する。

## 2. Goal
既存の `exstruct_patch` フローを維持しつつ、AIエージェントが少ない入力で安全にデザイン編集できるようにする。

## 3. In Scope
- グリッド罫線描画
- セル太字
- セル背景色
- 行高/列幅編集
- `return_inverse_ops` での新op逆操作返却

## 4. Out of Scope
- 詳細な罫線スタイル指定（MVPでは固定）
- グラデーション塗りなど高度塗り設定
- 新規ツール追加（`exstruct_patch` 拡張で対応）

## 5. Public Interface Changes
`PatchOp.op` に以下を追加する。
- `draw_grid_border`
- `set_bold`
- `set_fill_color`
- `set_dimensions`
- `restore_design_snapshot`（内部用）

### 5.1 op仕様
1. `draw_grid_border`
- 必須: `sheet`, `base_cell`, `row_count`, `col_count`
- 動作: `base_cell` 起点の矩形全セルに罫線を適用
- MVP仕様: `thin + black` 固定

2. `set_bold`
- 必須: `sheet` と (`cell` または `range`)
- 任意: `bold`（デフォルト `true`）
- 動作: 対象セルのフォント太字を設定

3. `set_fill_color`
- 必須: `sheet` と (`cell` または `range`) と `fill_color`
- 色形式: `#RRGGBB` または `#AARRGGBB`
- 動作: `solid` 塗りで背景色設定

4. `set_dimensions`
- 必須: `sheet`
- 行設定: `rows` + `row_height`
- 列設定: `columns` + `column_width`
- `columns` は列記号（A/AA）または数値の両対応
- 行/列は片方のみ、または両方同時指定可

5. `restore_design_snapshot`（内部用）
- `inverse_ops` で返した復元情報を再適用するためのop

## 6. Backend Policy
- 新デザインopは openpyxl 経路で処理する。
- 既存の値更新系COM経路は維持する。
- `.xls` でデザインopが含まれる場合は明示エラーとする。

## 7. Validation Rules
- `draw_grid_border`: `row_count >= 1`, `col_count >= 1`
- `set_bold`/`set_fill_color`: `cell` と `range` の同時指定禁止、未指定禁止
- `set_fill_color`: 色フォーマット厳格検証
- `set_dimensions`: 行/列指定と寸法値の組を必須化、寸法は正数
- 大規模誤操作防止: 対象セル数上限を設定（例: 10,000）

## 8. Diff / Undo
- `patch_diff` は既存形式を維持し、`kind` に style/dimension を追加
- `return_inverse_ops=true` 時は新opも逆操作を返す

## 9. Test Scenarios
- 各新opの正常系（単一セル、範囲、複合指定）
- 異常系（必須欠落、形式不正、競合指定、不正寸法）
- 逆操作の往復検証（適用→inverse再適用で復元）
- JSON文字列ops経由のサーバー層受理確認
- `.xls` 制約とエラーメッセージ確認

## 10. Acceptance Criteria
- 4機能が `exstruct_patch` で利用可能
- mypy strict / Ruff / tests が通過
- 既存op回帰なし
- MCPドキュメントとREADME更新済み

## 11. Assumptions / Defaults
- 罫線は MVP として固定スタイル
- 背景色は solid のみ
- `restore_design_snapshot` は内部用途だが受理可能
- 更新対象ドキュメントは `docs/agents/FEATURE_SPEC.md`
