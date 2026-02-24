# Task List

未完了: `[ ]` / 完了: `[x]`

## Epic: MCP Patch Architecture Refactor (Phase 1)

## 0. 事前準備と合意

- [x] `FEATURE_SPEC.md` と本タスクの整合性確認
- [x] 既存公開 API（import 経路・MCP I/F）の互換条件を明文化
- [x] 回帰対象テスト群の確定（patch/make/server/tools）

完了条件:
- [x] 仕様・互換条件・テスト対象がレビューで承認されている

## 1. 共通ユーティリティ抽出（低リスク先行）

- [x] `src/exstruct/mcp/shared/a1.py` を追加
- [x] A1/列変換関数を `patch_runner.py`・`server.py` から移設
- [x] `src/exstruct/mcp/shared/output_path.py` を追加
- [x] 出力 path 解決/競合処理を `patch_runner.py`・`extract_runner.py` から移設
- [x] 既存呼び出し元を共通ユーティリティ利用へ置換

完了条件:
- [x] A1 と output path の重複実装が削除されている
- [x] 関連テストが回帰なしで通る

## 2. patch ドメイン分離（型とモデル）

- [x] `src/exstruct/mcp/patch/types.py` を追加
- [x] `PatchOpType` ほか patch 共通型を移設
- [x] `src/exstruct/mcp/patch/models.py` を追加
- [x] `PatchOp` / `PatchRequest` / `MakeRequest` / `PatchResult` と snapshot モデルを移設
- [x] `patch_runner.py` から新モジュールを再エクスポート

完了条件:
- [x] モデルが `patch_runner.py` 以外からも直接利用可能
- [x] `patch_runner.py` のモデル定義が削減されている

## 3. 正規化と仕様メタデータの一元化

- [x] `src/exstruct/mcp/patch/specs.py` を追加
- [x] op ごとの required/optional/constraints/aliases を集約
- [x] `src/exstruct/mcp/patch/normalize.py` を追加
- [x] top-level `sheet` 解決と alias 正規化を移設
- [x] `server.py` の `_coerce_patch_ops` 系を共通ロジック利用へ置換
- [x] `tools.py` の top-level `sheet` 解決を共通ロジック利用へ置換
- [x] `op_schema.py` の `PatchOpType` 依存を `patch/specs.py` / `patch/types.py` へ変更

完了条件:
- [x] patch op 正規化実装が単一ソース化されている
- [x] `server.py` と `tools.py` の重複ロジックが削減されている

## 4. サービス層と backend 分離

- [x] `src/exstruct/mcp/patch/service.py` を追加
- [x] `run_patch` / `run_make` のオーケストレーションを移設
- [x] `src/exstruct/mcp/patch/engine/base.py` を追加（engine protocol）
- [x] `openpyxl` 実装を `engine/openpyxl_engine.py` へ移設
- [x] `xlwings` 実装を `engine/xlwings_engine.py` へ移設
- [x] 必要に応じて op 実装を `patch/ops/*` へ分離
- [x] `patch_runner.py` を薄いファサードへ縮退

完了条件:
- [x] `patch_runner.py` の主責務が公開互換維持のみになっている
- [x] engine 分岐/実装が `service.py` と `engine/*` に分離されている

## 5. テスト再配置と追加

- [x] `tests/mcp/patch/test_normalize.py` を追加
- [x] `tests/mcp/patch/test_service.py` を追加
- [x] `tests/mcp/shared/test_a1.py` を追加
- [x] `tests/mcp/shared/test_output_path.py` を追加
- [x] 既存テストを責務別に分割（必要箇所のみ）
- [x] `tests/mcp/test_patch_runner.py` の互換観点テストを維持

完了条件:
- [x] 新規分割モジュールに直接対応するテストが存在する
- [x] 既存互換テストが通る

## 6. ドキュメント更新

- [x] `docs/agents/ARCHITECTURE.md` に新構成を反映
- [x] `docs/agents/DATA_MODEL.md` の patch モデル参照先を更新
- [x] 必要に応じて `docs/mcp.md` の内部実装説明を更新

完了条件:
- [x] 参照先のコードパスが現行構成と一致している

## 7. 品質ゲート

- [x] `uv run task precommit-run` 実行
- [x] 失敗時は修正して再実行
- [x] 変更差分の自己レビュー（責務分離・循環依存・互換性）

完了条件:
- [x] mypy strict: 0 エラー
- [x] Ruff: 0 エラー
- [x] テスト: 全て成功

## 8. レガシー実装完全廃止（Phase 2）

- [x] `src/exstruct/mcp/patch/legacy_runner.py` 依存の棚卸し（import/呼び出し元を全列挙）
- [x] `patch/service.py` / `patch/engine/*` の `legacy_runner` 依存を新モジュール群へ置換
- [x] `patch/models.py` の `patch_runner` 経由 import を廃止し、実体モデル定義へ移行
- [x] `patch_runner.py` の monkeypatch 互換レイヤを段階的に削除（必要な公開 API は維持）
- [x] `tests/mcp/test_patch_runner.py` の私有関数前提テストを責務別テストへ移管
- [x] `src/exstruct/mcp/patch/ops/*` を導入し、op 実装を backend 別に分離
- [x] `legacy_runner.py` を削除し、不要な再エクスポートを整理
- [x] 互換性要件を満たしたまま `uv run task precommit-run` と回帰テストを再通過

完了条件:
- [x] `legacy_runner.py` がリポジトリから削除されている
- [x] `patch_runner.py` が公開 API の薄い入口のみを保持している
- [x] patch 実装の依存方向が `service -> engine/ops` に一本化されている
- [x] 既存 MCP I/F 互換とテスト成功が維持されている

## 優先順位

1. P0: 1, 2, 3
2. P1: 4, 5
3. P2: 6, 7, 8

## マイルストーン（推奨）

1. M1: 共通ユーティリティ抽出完了（Task 1）
2. M2: ドメイン/正規化分離完了（Task 2-3）
3. M3: service/engine 分離完了（Task 4）
4. M4: テスト・ドキュメント・品質ゲート完了（Task 5-7）
5. M5: レガシー実装完全廃止完了（Task 8）

---

## Epic: MCP Coverage Recovery (Post-Refactor)

## 0. 現状固定と差分計測

- [x] `coverage.xml` を基準値として保存（78.24% / miss 1,654）
- [x] 低下主因3ファイルの未実行行を記録（internal/models/server）
- [ ] 改善後比較用コマンドを固定化
- [ ] 完了条件: before/after の比較表が作成されている

## 1. `patch/models.py` 分岐網羅

- [x] `PatchOp` 各 validator の失敗系を `parametrize` で追加
- [x] alias競合・必須不足・型不正・範囲不正のケースを追加
- [x] `set_style` / `set_alignment` / `set_dimensions` の境界値ケースを追加
- [ ] 完了条件: `models.py` の未実行行を大幅削減（目安 80+ 行カバー）

## 2. `patch/internal.py` 分岐網羅

- [ ] openpyxl 適用系（成功/失敗/skip）を fixtureベースで追加
- [ ] `dry_run` / `preflight_formula_check` / `return_inverse_ops` の分岐を追加
- [x] unsupported op / sheet not found / range shape mismatch を網羅
- [ ] conflict policy（overwrite/rename/skip）の分岐を網羅
- [ ] 完了条件: `internal.py` の未実行行を大幅削減（目安 250+ 行カバー）

## 3. `server.py` 未カバー経路の補完

- [x] alias正規化 helper のエラー経路テストを追加
- [x] draw_grid_border range shorthand の不正入力テストを追加
- [x] patch op JSON parse の例外文言テストを追加
- [ ] 完了条件: `server.py` の line-rate を有意改善（目安 +10pt 以上）

## 4. Reader系の境界ケース補完

- [x] `test_sheet_reader.py` に invalid range / empty result / boundary を追加
- [x] `test_chunk_reader.py` に cursor/filter/max_bytes 境界テストを追加
- [ ] 完了条件: `sheet_reader.py` と `chunk_reader.py` の未実行行を削減

## 5. CIゲート強化

- [x] テストコマンドに `--cov-fail-under=80` を追加
- [x] `codecov.yml` の patch target を `80%` に設定
- [ ] PR時に project/patch 両ステータスを required として運用
- [ ] 完了条件: 80% 未満のPRがCIで確実に失敗する

## 6. 検証

- [x] `uv run task precommit-run` 実行
- [ ] `uv run pytest -m "not com and not render" --cov=exstruct --cov-report=xml --cov-fail-under=80` 実行
- [x] `uv run coverage report -m` で改善確認
- [ ] 完了条件: 全体80%以上、主要低下ファイルの改善、静的解析0エラー
