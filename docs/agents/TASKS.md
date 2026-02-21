# Task List

未完了 [ ], 完了 [x]

## Epic: MCP UX Hardening

### 0. 事前準備

- [ ] `sample.md` の失敗ケースをテストケース化する方針を確定する
- [ ] 既存の `tests/mcp/` で再利用可能な fixture を確認する

### 1. 入力正規化（FS-01）

- [x] `src/exstruct/mcp/server.py` の `_coerce_patch_ops` を拡張し、`name -> sheet` を追加
- [x] `src/exstruct/mcp/server.py` の `_coerce_patch_ops` を拡張し、`col/row -> columns/rows` を追加
- [x] `src/exstruct/mcp/server.py` の `_coerce_patch_ops` を拡張し、`width/height -> column_width/row_height` を追加
- [x] 正式フィールドとエイリアスが矛盾する場合のエラー仕様を実装
- [x] `color -> fill_color` 自動変換を無効化し、仕様として禁止
- [x] 上記変換の単体テストを追加（正常系/矛盾系）

完了条件:
- [x] `name` / `col` / `width` 入力が `PatchOp` 検証前に正規化される
- [x] `color` と `fill_color` が混同されない

### 2. HEX自由指定（FS-02）

- [x] `src/exstruct/mcp/patch_runner.py` の色バリデータを拡張し、`color`/`fill_color` ともに `#` なし 6/8 桁を許容
- [x] 内部正規化（`#` 補完 + 大文字化）を実装
- [x] 不正桁数・不正文字のエラーを維持
- [x] テスト追加（6桁/8桁/`#`あり/`#`なし/不正値）

完了条件:
- [x] 任意HEX指定が成功する
- [x] 不正入力は明確なエラーになる

### 3. 文字色 `set_font_color` 追加（FS-03）

- [x] `src/exstruct/mcp/patch_runner.py` の `PatchOpType` に `set_font_color` を追加
- [x] `PatchOp` に `color` フィールドを追加
- [x] `set_font_color` 専用バリデーションを追加（`sheet` + `color` + exactly one of `cell`/`range`）
- [x] `set_font_color` 実処理を openpyxl/com 両方に追加
- [x] `set_font_color` で `fill_color` 指定時にエラー化
- [x] テスト追加（単セル/範囲/エラーケース）

完了条件:
- [x] `color` は文字色として適用される
- [x] `fill_color` とは独立して動作する

### 4. draw_grid_border shorthand（FS-04）

- [x] `src/exstruct/mcp/server.py` または `src/exstruct/mcp/patch_runner.py` で `range` から `base_cell/row_count/col_count` へ正規化
- [x] `range` と `base_cell/row_count/col_count` の併用エラーを実装
- [x] 最大セル数制限の既存チェックが有効なことを確認
- [x] テスト追加（正常系/併用エラー/上限制約）

完了条件:
- [x] `draw_grid_border + range` が成功する
- [x] 併用ケースで明確なエラーになる

### 5. パス UX 改善（FS-05）

- [x] 相対 `out_path` 解決ルールを root 基準で統一
- [x] `Path is outside root` エラー文言に root と有効例を追加
- [x] テスト追加（相対成功/絶対失敗/診断メッセージ）

完了条件:
- [x] 相対 `out_path` が環境依存せず安定
- [x] root 外エラーの自己修復性が上がる

### 6. ランタイム情報ツール（FS-06, optional）

- [ ] `src/exstruct/mcp/server.py` に `exstruct_get_runtime_info` を追加
- [ ] `src/exstruct/mcp/tools.py` に対応入出力モデルを追加
- [ ] テスト追加（root/cwd/platform/path_examples の整合性）
- [ ] `docs/mcp.md` に利用例を追加

完了条件:
- [ ] エージェントが root/cwd を1コールで取得できる

### 7. ドキュメント整備

- [x] `docs/mcp.md` に「色指定の仕様（color / fill_color）」節を追加
- [x] `docs/mcp.md` に `set_font_color` / `set_fill_color` の最小例を追加
- [x] `docs/mcp.md` に `set_dimensions` / `draw_grid_border` の最小例を追加
- [ ] `docs/README.ja.md` の MCP 節に注意点を反映

完了条件:
- [ ] 同種ミスの再発を抑制できるドキュメントになっている

### 8. 検証・リリース

- [x] `uv run task precommit-run` を実行し、mypy / Ruff / pytest を通す
- [ ] 変更点を `docs/release-notes/` に追加
- [ ] PR テンプレートに受け入れ条件のチェックを記載

完了条件:
- [ ] CI グリーン
- [ ] Feature Spec の AC-01〜AC-06 を満たす

## 優先順位

1. P0: 1, 2, 3, 4, 5
2. P1: 7
3. P2: 6, 8
