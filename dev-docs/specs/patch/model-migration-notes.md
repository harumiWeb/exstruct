# Patch Model 移行メモ（Phase 2）

`patch/models.py` を canonical model として一本化する際の依存メモです。

## 現状の結合点

- `patch/models.py` に `PatchOp` / `PatchRequest` / `PatchResult` などの canonical 定義があり、`internal.py` に重複定義が残っている
- `patch_runner.py` / `service.py` / `runtime.py` / `ops/*` は `internal.py` の private 実装に依存している
- そのため、両系統の `BaseModel` が混在すると、mypy と実行時検証の両方で型不整合が発生

## 段階移行の推奨手順

1. `internal.py` の重複モデル定義を削除し、`patch/models.py` からの import に置き換える
2. `internal.py` 内のモデルバリデーション補助関数（`PatchOp` 関連）を `models.py` 側に移設
3. `runtime.py` / `ops/*` / `service.py` の型注釈と返却型を `patch.models` へ統一
4. `tests/mcp/test_patch_runner.py` の互換テストを維持したまま `internal.py` 依存テストを `tests/mcp/patch/*` へ移管
5. 最後に `internal.py` への互換依存を縮退し、`patch_runner.py` を公開 API の薄い入口に固定

## 注意点

- `internal.py` の重複定義を残したまま呼び出し側を切り替えると、`PatchResult` 構築時に異なるクラス階層が混在して検証に失敗しやすい
- まずは **定義元を一本化** してから呼び出し側を差し替えるのが安全
