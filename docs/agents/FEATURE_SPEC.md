# Feature Spec for AI Agent

## Feature Name

MCP Patch Architecture Refactor (Phase 1)

## 背景

`src/exstruct/mcp/patch_runner.py` は 3,500 行超の単一モジュールとなっており、以下が混在している。

1. ドメインモデル定義（`PatchOp`, `PatchRequest`, `PatchResult`）
2. 入力検証（op 別バリデーション）
3. 実行制御（engine 選択、fallback、warning 集約）
4. backend 実装（openpyxl/xlwings）
5. 共通ユーティリティ（A1 変換、path 競合処理、色変換）

この状態は、保守性・拡張性・テスト容易性を低下させるため、責務分離を行う。

## 目的

1. `patch_runner.py` を薄いファサードへ縮退する。
2. patch 機能を「モデル」「正規化/検証」「実行制御」「backend 実装」に分離する。
3. `server.py` と `patch_runner.py` に分散した patch op 正規化を共通化する。
4. 重複ユーティリティ（A1、出力 path 競合処理）を共通化する。
5. 公開 API 互換を維持しつつ、段階的に移行可能な構造にする。

## スコープ

### In Scope

1. `src/exstruct/mcp/patch/` 配下の新規モジュール導入
2. `patch_runner.py` のファサード化
3. patch op 正規化の共通化
4. A1 変換・出力 path 解決の共通ユーティリティ化
5. 既存テストの責務別再配置と追加

### Out of Scope

1. 新しい patch op の追加
2. MCP 外部 API の仕様変更
3. 大規模なディレクトリ再編（`mcp` 全体の再設計）
4. `.xls` サポート方針の変更

## ターゲットアーキテクチャ

```text
src/exstruct/mcp/
  patch/
    __init__.py
    types.py
    models.py
    specs.py
    normalize.py
    validate.py
    service.py
    engine/
      base.py
      openpyxl_engine.py
      xlwings_engine.py
    ops/
      common.py
      openpyxl_ops.py
      xlwings_ops.py
  shared/
    a1.py
    output_path.py
```

## モジュール責務

1. `patch/types.py`
   1. `PatchOpType`, `PatchBackend`, `PatchEngine`, `PatchStatus` 等の型定義
2. `patch/models.py`
   1. `PatchOp`, `PatchRequest`, `MakeRequest`, `PatchResult` と snapshot 系モデル
3. `patch/specs.py`
   1. op ごとの required/optional/constraints/aliases を単一管理
4. `patch/normalize.py`
   1. top-level `sheet` 適用
   2. alias 正規化（`name`, `row`, `col`, `horizontal`, `vertical`, `color` など）
5. `patch/validate.py`
   1. `PatchOp` の整合性チェック（spec ベース）
6. `patch/service.py`
   1. `run_patch` / `run_make` のオーケストレーション
   2. engine 選択・fallback・warning/error 組み立て
7. `patch/engine/*`
   1. backend ごとの workbook 編集と保存責務
8. `patch/ops/*`
   1. op 適用ロジック（backend 別）
9. `shared/a1.py`
   1. A1 解析、列変換、範囲展開の共通関数
10. `shared/output_path.py`
   1. `on_conflict`、`rename`、出力先決定の共通関数

## 依存ルール

1. `server.py` は patch 実装詳細に依存しない。
2. `op_schema.py` は `patch_runner.py` ではなく `patch/specs.py` / `patch/types.py` に依存する。
3. `tools.py` は `patch/service.py` と `patch/models.py` のみを利用する。
4. backend 実装は `service.py` への逆依存を禁止する。
5. 共通関数は `shared/*` に集約し、重複実装を禁止する。

## 互換性要件

1. 既存の公開 import は維持する（`exstruct.mcp.patch_runner` 経由の主要シンボル）。
2. MCP tool I/F は変更しない（入力・出力 JSON 互換）。
3. 既存 warning/error メッセージは可能な限り維持する。
4. 既存テスト（`tests/mcp/test_patch_runner.py` ほか）を通す。

## 非機能要件

1. mypy strict: エラー 0
2. Ruff (E, W, F, I, B, UP, N, C90): エラー 0
3. 循環依存 0
4. 1 モジュール 1 責務を優先し、巨大関数を分割する
5. 新規関数/クラスは Google スタイル docstring を付与する

## 受け入れ条件（Acceptance Criteria）

### AC-01 モジュール分離

1. `patch_runner.py` がファサード化され、実装詳細の大半が `patch/` 配下へ移動している。

### AC-02 正規化一元化

1. patch op alias 正規化ロジックが単一モジュールに集約され、`server.py` と `tools.py` から再利用される。

### AC-03 重複削減

1. A1 変換の重複実装が除去され、`shared/a1.py` に統一される。
2. 出力 path 競合処理の重複実装が除去され、`shared/output_path.py` に統一される。

### AC-04 互換性維持

1. 既存の MCP ツール呼び出しが回帰なく動作する。
2. 既存 patch/make 関連テストが通過する。

### AC-05 品質ゲート

1. `uv run task precommit-run` が成功する。

## テスト方針

1. 既存テストの回帰確認
   1. `tests/mcp/test_patch_runner.py`
   2. `tests/mcp/test_make_runner.py`
   3. `tests/mcp/test_server.py`
   4. `tests/mcp/test_tools_handlers.py`
2. 新規テスト追加
   1. `tests/mcp/patch/test_normalize.py`
   2. `tests/mcp/patch/test_service.py`
   3. `tests/mcp/shared/test_a1.py`
   4. `tests/mcp/shared/test_output_path.py`

## リスクと対策

1. リスク: 分割中に import 互換が崩れる
   1. 対策: `patch_runner.py` で再エクスポートを維持し、段階移行する
2. リスク: warning/error 文言差分でテストが壊れる
   1. 対策: 既存文言互換を維持し、必要時は差分を明示してテスト更新する
3. リスク: engine 分離時の挙動差
   1. 対策: backend ごとの回帰テストを先に固定してから移行する
