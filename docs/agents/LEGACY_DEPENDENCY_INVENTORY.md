# Legacy Dependency Inventory (Phase 2)

更新日: 2026-02-24

`src/exstruct/mcp/patch/legacy_runner.py` は Phase 2 完了時に削除済みです。
このドキュメントは、旧依存の棚卸し結果と現行の置換先を記録します。

## 旧依存の置換先

- 旧対象: `src/exstruct/mcp/patch/legacy_runner.py`（削除済み）
- 現行の責務分割:
  - `src/exstruct/mcp/patch/service.py`: patch/make のオーケストレーション
  - `src/exstruct/mcp/patch/engine/openpyxl_engine.py`: openpyxl backend 実行境界
  - `src/exstruct/mcp/patch/engine/xlwings_engine.py`: COM(xlwings) backend 実行境界
  - `src/exstruct/mcp/patch/runtime.py`: runtime ユーティリティ（engine選択・path・policy）
  - `src/exstruct/mcp/patch/ops/*`: backend 別 op 適用ロジック

## 互換レイヤ

- `src/exstruct/mcp/patch_runner.py`
  - 公開 import 互換を維持する薄いファサード
  - 実体実装は `patch/service.py` 側に委譲

## テスト観点

- `tests/mcp/test_patch_runner.py`
  - `patch_runner` の互換入口（委譲動作）を検証
- `tests/mcp/patch/test_service.py`
  - backend 選択・fallback・警告メッセージを検証
