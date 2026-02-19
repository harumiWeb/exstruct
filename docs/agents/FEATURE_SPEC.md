# Feature Spec for AI Agent (Phase-by-Phase)

## 1. Feature
Issue: MCP patch operation for font size control (`set_font_size`)

## 2. Goal
`exstruct_patch` で文字サイズを明示指定できるようにし、見出し・本文・注記の視覚階層を安定して表現できるようにする。

## 3. In Scope
- 新規op `set_font_size` の追加
- `cell` または `range` を対象に `font_size` を適用
- openpyxl / COM の両経路で適用
- 既存 `patch_diff` 互換を維持
- `exstruct_patch` docstring / docs 追記

## 4. Out of Scope
- フォント名変更（`font_name`）
- フォント色変更（`font_color`）
- 斜体/下線など追加装飾op
- 既存op名・既存スキーマの破壊的変更

## 5. Public Interface Changes

### 5.1 `PatchOp`
- `op="set_font_size"` を追加
- 必須:
  - `sheet: str`
  - `font_size: float`
- 対象指定:
  - `cell` または `range` のどちらか一方を必須

### 5.2 `server.py` (`exstruct_patch` docstring)
- 対応op一覧に `set_font_size` を追加
- 引数仕様（範囲・制約）を明記

## 6. Validation Rules
- `set_font_size` は `cell` / `range` を同時指定不可、未指定不可
- `font_size` は `> 0` 必須
- `set_font_size` で無関係フィールド（`value`, `formula`, `fill_color`, `rows`, `columns` など）は受け付けない
- スタイル対象セル数の上限チェックは既存スタイルopと同等ルールを適用

## 7. Backend Behavior

### 7.1 openpyxl
- 対象セルの既存フォント属性（bold/italic/name 等）を維持しつつ `size` のみ更新

### 7.2 COM(xlwings)
- 対象レンジの `Font.Size` を更新
- `patch_diff` は既存スタイルopに準じた表現で返却

## 8. Compatibility
- 既存opの挙動は不変
- `PatchResult` スキーマ変更なし
- `patch_diff` の構造は既存形式を維持

## 9. Test Scenarios
- `set_font_size` の正常系（cell指定）
- `set_font_size` の正常系（range指定）
- `font_size <= 0` の異常系
- `cell` と `range` 同時指定の異常系
- `cell` / `range` 未指定の異常系
- openpyxl経路で他フォント属性保持を検証
- COM経路で適用されることを検証
- 既存op回帰（影響なし）

## 10. Acceptance Criteria
- `set_font_size` で指定セル/範囲の文字サイズが更新される
- バリデーション異常系が期待通り失敗する
- `uv run pytest tests/mcp` が通過
- `uv run task precommit-run` が通過

## 11. Assumptions / Defaults
- まずは `font_size` のみ追加し、他のフォント属性は追加しない
- 更新対象ドキュメントは `docs/agents/FEATURE_SPEC.md` と `docs/agents/TASKS.md`
