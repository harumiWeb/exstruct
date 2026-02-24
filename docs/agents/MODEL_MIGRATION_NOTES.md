# Patch Model Migration Notes (Phase 2)

`patch/models.py` の実体モデル化を進める際の依存メモです。

## 現状の結合点

- `legacy_runner.py` が `PatchOp` / `PatchRequest` / `PatchResult` などの実体定義を保持
- `service.py` / `runtime.py` / `ops/*` は `legacy_runner` の private 関数を呼び出す
- そのため、`models.py` に同名の別 `BaseModel` を作ると、mypy と実行時検証の両方で型不整合が発生

## 段階移行の推奨手順

1. `legacy_runner.py` のモデル定義を `patch/models.py` に移し、`legacy_runner.py` では import のみを行う
2. `legacy_runner.py` 内のモデルバリデーション補助関数（`PatchOp` 関連）を `models.py` 側に移設
3. `runtime.py` / `ops/*` / `service.py` の型注釈を `patch.models` へ統一
4. `tests/mcp/test_patch_runner.py` の互換テストを維持したまま `legacy_runner.py` 依存テストを `tests/mcp/patch/*` へ移管
5. 最後に `legacy_runner.py` を削除し、`patch_runner.py` を公開 API の薄い入口に固定

## 注意点

- `PatchDiffItem` / `PatchErrorDetail` / `FormulaIssue` を先行して別モデル化すると、`PatchResult` 構築時の検証で失敗しやすい
- まずは **定義元を一本化** してから呼び出し側を差し替えるのが安全
