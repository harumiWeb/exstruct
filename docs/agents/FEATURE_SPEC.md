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

---

## Feature Name

MCP Coverage Recovery (Post-Refactor)

## 背景

MCP 大規模リファクタリング後、全体カバレッジが 80% から 78.24% に低下した。
`coverage.xml` の未実行行は 1,654 行で、うち `src/exstruct/mcp/*` が 1,176 行（71.1%）を占める。

主因:
1. `src/exstruct/mcp/patch/internal.py`（59.42%, 806 miss）
2. `src/exstruct/mcp/patch/models.py`（77.74%, 181 miss）
3. `src/exstruct/mcp/server.py`（70.17%, 71 miss）

## 目的

1. 全体カバレッジを 80%以上へ回復し維持する。
2. 低下要因モジュールをテストで直接改善する。
3. `omit` 依存の見かけ上の回復は行わない。

## スコープ

### In Scope

1. `tests/mcp/patch/*` の拡張（models/internal/service中心）
2. `tests/mcp/test_server.py` の未カバー分岐追加
3. `tests/mcp/test_sheet_reader.py` / `tests/mcp/test_chunk_reader.py` の境界ケース追加
4. CI ゲート設定の強化（`--cov-fail-under=80` と patch coverage 80%）
5. 追記ドキュメント（本仕様・タスク・テスト要件）

### Out of Scope

1. patch op の新規機能追加
2. 公開 API 仕様変更
3. 大規模ディレクトリ再編

## 実装方針

1. `patch/models.py` の validator 分岐を `pytest.mark.parametrize` で網羅する。
2. `patch/internal.py` の openpyxl/xlwings 分岐、エラー分岐、保存可否分岐を網羅する。
3. `server.py` の alias 正規化・A1 パース・エラーメッセージ経路を網羅する。
4. `sheet_reader.py` / `chunk_reader.py` の境界ケースを追加する。
5. CI を「全体80%未満で失敗」「変更行80%未満で失敗」にする。

## 公開API/インターフェース変更

1. Python 公開 API の変更は行わない。
2. CI インターフェースとして以下を追加・変更する。
3. テスト実行コマンドに `--cov-fail-under=80` を追加する。
4. Codecov `patch` ステータス目標を `80%` に設定する。

## 受け入れ基準（Acceptance Criteria）

1. `uv run pytest -m "not com and not render" --cov=exstruct --cov-report=xml --cov-fail-under=80` が成功する。
2. `coverage.xml` の全体 line-rate が 80%以上である。
3. Codecov patch coverage の required status が 80%以上である。
4. `patch/internal.py`, `patch/models.py`, `server.py` の line-rate が現状より改善している。
5. `uv run task precommit-run` が成功する（mypy strict / Ruff 含む）。

## リスクと対策

1. リスク: `patch/internal.py` の分岐が多く工数が膨らむ。
   対策: 失敗系を `parametrize` 化し、網羅効率を最大化する。
2. リスク: CI ゲート強化で一時的に失敗が増える。
   対策: 不足テストを同PRで同時投入する。
3. リスク: 互換レイヤー除外に後退する。
   対策: 恒久除外を禁止し、必要時は別承認で期限付き措置とする。

---

## Feature Name

PR #65 Review Follow-up (Stabilization)

## 背景

2026-02-24 時点で PR #65 には GitHub 上で未解決レビュー指摘が残っている。
主要な論点は次の 3 系統。

1. 機能不具合: `backend="auto"` かつ `apply_table_style` で COM が選択されると失敗する。
2. ドキュメント不整合: `legacy_runner` 参照や alias 許容/禁止の記述が実装と一致しない。
3. 設計改善提案: BaseModel 境界、tuple/dict 返却、Protocol 整理、A1戻り値モデル化など。

## 目的

1. マージ阻害となる不具合・矛盾（P0）を優先解消する。
2. 低リスクで回収できる品質改善（P1）を同PRで解消する。
3. 変更波及の大きい設計変更（P2）は別Epicへ分離し、短期安定性を優先する。

## 対応方針（指摘分類）

### P0: 同PRで必ず対応

1. `src/exstruct/mcp/patch/service.py`
   1. `apply_table_style` を含む場合、`backend="auto"` でも openpyxl 側へルーティングする。
2. `docs/agents/LEGACY_DEPENDENCY_INVENTORY.md`
   1. 削除済み `legacy_runner.py` 前提の記述を現行構成へ修正する。
3. `docs/mcp.md`
   1. Mistake catalog と alias 仕様の矛盾（`color`, `horizontal`, `vertical`）を解消する。

### P1: 同PRで実施する低リスク改善

1. `src/exstruct/mcp/server.py`
   1. 未使用・重複の正規化 helper 群を削除し、`patch.normalize` への一元化を明確化する。
2. Docstring 指摘
   1. 変更差分内で欠落している docstring を追加し、品質ゲートの指摘ノイズを減らす。
3. `src/exstruct/mcp/server.py::_register_tools`
   1. 追加済み引数（`default_on_conflict`, `artifact_bridge_dir`）を docstring に反映する。

### P2: 別Epicへ分離（今回は設計判断のみ）

1. `patch/__init__.py` 公開 API 方針
   1. 正規化 helper を公開し続けるか、内部化するかを決定する。
2. engine 戻り値の BaseModel 化（tuple 返却廃止）。
3. `shared/a1.py` 戻り値の BaseModel 化と呼び出し側移行。
4. `patch/internal.py` の外部ライブラリ境界（`Any` 受け + 内部正規化）再設計。

## スコープ

### In Scope

1. P0 の不具合修正とドキュメント整合。
2. P1 の低リスクなコード整理・docstring 修正。
3. P2 の設計判断結果をタスク化し、別Epicへ登録。

### Out of Scope

1. P2 の大規模設計変更の実装完了。
2. 公開 API の破壊的変更。
3. リリース範囲を拡大する新機能追加。

## 受け入れ基準（Acceptance Criteria）

1. `backend="auto"` + `apply_table_style` が Windows/COM 環境でも失敗しない。
2. `LEGACY_DEPENDENCY_INVENTORY.md` と `docs/mcp.md` の記述が現行実装と一致する。
3. PR #65 の P0 指摘が GitHub 上で解消済みになる。
4. P1 指摘は対応完了、または理由付きで deferred が明記される。
5. `uv run task precommit-run` が成功する。

## リスクと対策

1. リスク: P1 範囲が膨張し、P0 対応が遅延する。
   1. 対策: P0 完了前は P2 由来変更を着手しない。
2. リスク: 設計変更を同PRに混在させ、回帰リスクが上がる。
   1. 対策: P2 は issue 化して別PRへ分離する。
3. リスク: docstring 修正が大量化してレビュー負荷が上がる。
   1. 対策: 変更差分に限定して優先対応し、残件は別タスクへ切り出す。
