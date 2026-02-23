# Task List

未完了: `[ ]` / 完了: `[x]`

## Epic: MCP Patch Architecture Refactor (Phase 1)

## 0. 事前準備と合意

- [ ] `FEATURE_SPEC.md` と本タスクの整合性確認
- [ ] 既存公開 API（import 経路・MCP I/F）の互換条件を明文化
- [ ] 回帰対象テスト群の確定（patch/make/server/tools）

完了条件:
- [ ] 仕様・互換条件・テスト対象がレビューで承認されている

## 1. 共通ユーティリティ抽出（低リスク先行）

- [ ] `src/exstruct/mcp/shared/a1.py` を追加
- [ ] A1/列変換関数を `patch_runner.py`・`server.py` から移設
- [ ] `src/exstruct/mcp/shared/output_path.py` を追加
- [ ] 出力 path 解決/競合処理を `patch_runner.py`・`extract_runner.py` から移設
- [ ] 既存呼び出し元を共通ユーティリティ利用へ置換

完了条件:
- [ ] A1 と output path の重複実装が削除されている
- [ ] 関連テストが回帰なしで通る

## 2. patch ドメイン分離（型とモデル）

- [ ] `src/exstruct/mcp/patch/types.py` を追加
- [ ] `PatchOpType` ほか patch 共通型を移設
- [ ] `src/exstruct/mcp/patch/models.py` を追加
- [ ] `PatchOp` / `PatchRequest` / `MakeRequest` / `PatchResult` と snapshot モデルを移設
- [ ] `patch_runner.py` から新モジュールを再エクスポート

完了条件:
- [ ] モデルが `patch_runner.py` 以外からも直接利用可能
- [ ] `patch_runner.py` のモデル定義が削減されている

## 3. 正規化と仕様メタデータの一元化

- [ ] `src/exstruct/mcp/patch/specs.py` を追加
- [ ] op ごとの required/optional/constraints/aliases を集約
- [ ] `src/exstruct/mcp/patch/normalize.py` を追加
- [ ] top-level `sheet` 解決と alias 正規化を移設
- [ ] `server.py` の `_coerce_patch_ops` 系を共通ロジック利用へ置換
- [ ] `tools.py` の top-level `sheet` 解決を共通ロジック利用へ置換
- [ ] `op_schema.py` の `PatchOpType` 依存を `patch/specs.py` / `patch/types.py` へ変更

完了条件:
- [ ] patch op 正規化実装が単一ソース化されている
- [ ] `server.py` と `tools.py` の重複ロジックが削減されている

## 4. サービス層と backend 分離

- [ ] `src/exstruct/mcp/patch/service.py` を追加
- [ ] `run_patch` / `run_make` のオーケストレーションを移設
- [ ] `src/exstruct/mcp/patch/engine/base.py` を追加（engine protocol）
- [ ] `openpyxl` 実装を `engine/openpyxl_engine.py` へ移設
- [ ] `xlwings` 実装を `engine/xlwings_engine.py` へ移設
- [ ] 必要に応じて op 実装を `patch/ops/*` へ分離
- [ ] `patch_runner.py` を薄いファサードへ縮退

完了条件:
- [ ] `patch_runner.py` の主責務が公開互換維持のみになっている
- [ ] engine 分岐/実装が `service.py` と `engine/*` に分離されている

## 5. テスト再配置と追加

- [ ] `tests/mcp/patch/test_normalize.py` を追加
- [ ] `tests/mcp/patch/test_service.py` を追加
- [ ] `tests/mcp/shared/test_a1.py` を追加
- [ ] `tests/mcp/shared/test_output_path.py` を追加
- [ ] 既存テストを責務別に分割（必要箇所のみ）
- [ ] `tests/mcp/test_patch_runner.py` の互換観点テストを維持

完了条件:
- [ ] 新規分割モジュールに直接対応するテストが存在する
- [ ] 既存互換テストが通る

## 6. ドキュメント更新

- [ ] `docs/agents/ARCHITECTURE.md` に新構成を反映
- [ ] `docs/agents/DATA_MODEL.md` の patch モデル参照先を更新
- [ ] 必要に応じて `docs/mcp.md` の内部実装説明を更新

完了条件:
- [ ] 参照先のコードパスが現行構成と一致している

## 7. 品質ゲート

- [ ] `uv run task precommit-run` 実行
- [ ] 失敗時は修正して再実行
- [ ] 変更差分の自己レビュー（責務分離・循環依存・互換性）

完了条件:
- [ ] mypy strict: 0 エラー
- [ ] Ruff: 0 エラー
- [ ] テスト: 全て成功

## 優先順位

1. P0: 1, 2, 3
2. P1: 4, 5
3. P2: 6, 7

## マイルストーン（推奨）

1. M1: 共通ユーティリティ抽出完了（Task 1）
2. M2: ドメイン/正規化分離完了（Task 2-3）
3. M3: service/engine 分離完了（Task 4）
4. M4: テスト・ドキュメント・品質ゲート完了（Task 5-7）
