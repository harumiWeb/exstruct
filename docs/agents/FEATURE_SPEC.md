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
- 結合セル（結合/解除）
- 文字配置（水平/垂直/折返し）
- `return_inverse_ops` での新op逆操作返却

## 4. Out of Scope
- 詳細な罫線スタイル指定（MVPでは固定）
- グラデーション塗りなど高度塗り設定
- 回転、縮小表示、インデント、reading order などの高度配置
- 新規ツール追加（`exstruct_patch` 拡張で対応）

## 5. Public Interface Changes
`PatchOp.op` に以下を追加する。
- `draw_grid_border`
- `set_bold`
- `set_fill_color`
- `set_dimensions`
- `merge_cells`
- `unmerge_cells`
- `set_alignment`
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

5. `merge_cells`
- 必須: `sheet`, `range`
- 動作: 指定矩形を結合
- 制約: 既存結合範囲と交差した場合はエラー
- 注記: 左上以外に値がある場合は warning を返しつつ継続

6. `unmerge_cells`
- 必須: `sheet`, `range`
- 動作: 指定範囲に交差する結合範囲をすべて解除（該当なしは no-op）

7. `set_alignment`
- 必須: `sheet` と (`cell` または `range`)
- 必須: `horizontal_align` / `vertical_align` / `wrap_text` のうち少なくとも1つ
- 動作: 指定プロパティのみ更新（未指定プロパティは保持）
- 値制約:
  - `horizontal_align`: `general|left|center|right|fill|justify|centerContinuous|distributed`
  - `vertical_align`: `top|center|bottom|justify|distributed`

8. `restore_design_snapshot`（内部用）
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
- `merge_cells`: `range` 必須、単一セル範囲禁止、既存結合範囲交差禁止
- `unmerge_cells`: `range` 必須
- `set_alignment`: `cell/range` の片方必須、配置3項目のうち1つ以上必須
- 大規模誤操作防止: 対象セル数上限を設定（例: 10,000）

## 8. Diff / Undo
- `patch_diff` は既存形式を維持し、`kind` に style/dimension を追加
- `return_inverse_ops=true` 時は新opも逆操作を返す
- merge/alignment は `restore_design_snapshot` によるスナップショット復元で往復可能にする

## 9. Test Scenarios
- 各新opの正常系（単一セル、範囲、複合指定）
- 異常系（必須欠落、形式不正、競合指定、不正寸法）
- 逆操作の往復検証（適用→inverse再適用で復元）
- JSON文字列ops経由のサーバー層受理確認
- `.xls` 制約とエラーメッセージ確認
- merge時の値消失 warning 継続確認
- 既存結合との交差エラー確認
- set_alignment の未指定プロパティ保持確認

## 10. Acceptance Criteria
- 7機能が `exstruct_patch` で利用可能
- mypy strict / Ruff / tests が通過
- 既存op回帰なし
- MCPドキュメントとREADME更新済み

## 11. Assumptions / Defaults
- 罫線は MVP として固定スタイル
- 背景色は solid のみ
- 文字配置MVPは `horizontal_align` / `vertical_align` / `wrap_text` のみ
- 結合で失われる可能性のある値は warning で通知し、処理は継続
- `restore_design_snapshot` は内部用途だが受理可能
- 更新対象ドキュメントは `docs/agents/FEATURE_SPEC.md`
