# Legacy Dependency Inventory (Phase 2)

`src/exstruct/mcp/patch/legacy_runner.py` 依存の棚卸し結果です（2026-02-24）。

## 直接依存（コード）

- `src/exstruct/mcp/patch/runtime.py`
  - 理由: 既存 private 実装を互換維持しつつ段階移行するための集約レイヤ
- `src/exstruct/mcp/patch_runner.py`
  - 理由: 既存公開 import 経路・monkeypatch 互換の維持

## 間接依存（runtime 経由）

- `src/exstruct/mcp/patch/service.py`
- `src/exstruct/mcp/patch/engine/openpyxl_engine.py`
- `src/exstruct/mcp/patch/engine/xlwings_engine.py`

## テスト依存

- `tests/mcp/test_patch_runner.py`
  - 理由: `patch_runner` の私有関数 monkeypatch 互換を前提にしたテストが存在
