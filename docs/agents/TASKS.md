# Task List

未完了 [ ], 完了 [x]

## Epic: MCP UX Hardening Phase 2 (Claude Review Closure)

### 0. 仕様固定と設計（FS全体）

- [ ] `review.md` の指摘を FS/AC にマッピングし、非対象を明文化する
- [ ] 公開I/F変更一覧（型、ツール、CLI、レスポンス）を確定する
- [ ] 既存互換ポリシー（後方互換・原子性維持）を設計メモに固定する

完了条件:
- [ ] Feature Spec と実装対象が1対1で追跡できる

### 1. Validation UX（FS-01）

- [x] `_coerce_patch_ops` に alias 正規化を追加（`horizontal`/`vertical`/`color`）
- [x] `PatchErrorDetail` に `hint` / `expected_fields` / `example_op` を追加
- [x] 既知の入力ミスへ具体的ヒントを返すエラーヒント生成器を実装
- [x] 既存エラー経路（`PatchOpError.from_op` など）へ拡張項目を接続
- [x] テスト追加（alias正規化、ヒント内容、後方互換）

完了条件:
- [ ] 誤入力時に自己修復可能なエラー情報が返る

### 2. `set_style`（FS-02）

- [x] `PatchOpType` に `set_style` を追加
- [x] `PatchOp` validator を追加（target exactly one、属性1つ以上）
- [x] openpyxl 実装を追加（font/fill/alignment の複合適用）
- [x] com 実装を追加（同等の複合適用）
- [x] inverse snapshot（font/fill/alignment）復元を実装
- [x] テスト追加（単セル、範囲、属性未指定、上限超過）

完了条件:
- [ ] 1op で複数書式属性を安定適用できる

### 3. `apply_table_style`（FS-03）

- [x] `PatchOpType` に `apply_table_style` を追加
- [x] `PatchOp` に `style` / `table_name` を追加
- [x] validator を追加（必須項目、範囲妥当性、交差チェック前提）
- [x] openpyxl 実装を追加（Table + TableStyleInfo 適用）
- [x] com 指定時の warning + openpyxl フォールバック方針を実装
- [x] テスト追加（正常系、重複名、交差範囲、backend方針）

完了条件:
- [ ] テーブルスタイルを1opで適用できる

### 4. 成果物ミラー（FS-04）

- [x] `ServerConfig` / CLI に `--artifact-bridge-dir` を追加
- [x] `PatchToolInput` / `MakeToolInput` に `mirror_artifact` を追加
- [x] `PatchToolOutput` / `MakeToolOutput` に `mirrored_out_path` を追加
- [x] 成功時ミラーコピー処理を実装（bridge有効時のみ）
- [x] コピー失敗時は warning のみ返し、処理失敗にしない
- [x] テスト追加（正常、bridge未設定、コピー失敗）

完了条件:
- [ ] `present_files` 連携向けの成果物パスが返せる

### 5. 分割実行API（FS-05）

- [ ] `exstruct_patch_plan` の入出力モデルを追加
- [ ] `exstruct_patch_apply_chunks` の入出力モデルを追加
- [ ] サーバーへ新規2ツールを登録
- [ ] chunk生成ロジックを実装（`chunk_by`, `max_ops_per_chunk`）
- [ ] 内部ステージング + 原子コミット実装（成功時のみ最終保存）
- [ ] 失敗時ロールバック保証を実装（最終成果物未生成）
- [ ] テスト追加（計画妥当性、成功時、失敗時）

完了条件:
- [ ] 大量操作時も原子性を維持した分割実行が可能

### 6. 入力スキーマ可視化（FS-06）

- [ ] `exstruct_patch` ツール定義に `op` 別ミニスキーマ（required/optional/constraints/example）を追加
- [ ] `op` 別ミニスキーマに alias 対応を明記
- [ ] スキーマ記述の生成元を共通メタデータ化し、定義ドリフトを防止
- [ ] `exstruct_list_ops` の入出力モデルとサーバー登録を追加
- [ ] `exstruct_describe_op` の入出力モデルとサーバー登録を追加
- [ ] `exstruct_describe_op` に `required` / `optional` / `constraints` / `example` / `aliases` を実装
- [ ] テスト追加（一覧妥当性、describe内容、未知opエラー、tool定義文言）

完了条件:
- [ ] 実行前に主要 `op` の入力仕様と動作例を確認できる

### 7. ドキュメント整備

- [ ] `docs/mcp.md` に `set_style` を追記（最小例、制約、エラー例）
- [ ] `docs/mcp.md` に `apply_table_style` を追記（最小例、制約）
- [ ] `docs/mcp.md` に `mirror_artifact` / `mirrored_out_path` を追記
- [ ] `docs/mcp.md` に `exstruct_patch_plan` / `exstruct_patch_apply_chunks` を追記
- [ ] `docs/mcp.md` に `exstruct_list_ops` / `exstruct_describe_op` を追記
- [ ] 失敗例→正解例（引数名ミス）カタログを追記

完了条件:
- [ ] レビューで指摘された試行錯誤パターンをドキュメントで回避できる

### 8. 検証・受け入れ

- [ ] `uv run task precommit-run` を実行
- [ ] 既存回帰テスト + 新規ACテストが通過
- [ ] AC-01 〜 AC-07 の達成をチェックリストで確認

完了条件:
- [ ] CI グリーン
- [ ] Feature Spec の受け入れ条件を満たす

## 優先順位

1. P0: 1, 2, 3, 4, 5, 6（ただし6は「ツール定義拡充」を最優先）
2. P1: 6（`exstruct_list_ops` / `exstruct_describe_op`）, 7
3. P2: 8

## テストケース（必須追跡）

- [x] パラメータ誤り時のヒント返却（`color` vs `fill_color`、`horizontal` vs `horizontal_align`）
- [x] `set_style` の単セル/範囲/属性未指定エラー
- [x] `apply_table_style` の正常系/重複テーブルエラー
- [x] `mirror_artifact` の正常コピー/bridge未設定/コピー失敗warning
- [ ] `patch_plan` のchunk生成妥当性
- [ ] `patch_apply_chunks` の成功時コミット、失敗時ロールバック
- [ ] `exstruct_list_ops` の一覧妥当性
- [ ] `exstruct_describe_op` の required/optional/example 妥当性
- [ ] `exstruct_patch` ツール定義に `op` 別スキーマ情報が含まれること
