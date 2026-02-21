# Feature Spec for AI Agent (Phase-by-Phase)

## Feature Name

MCP UX Hardening for `exstruct_make` / `exstruct_patch`

## 背景

`sample.md` の会話ログでは、`exstruct_make` 呼び出し 10 回中 6 回が失敗しており、主因は以下です。

1. 操作パラメータ名の不一致（例: `col` / `width` / `name`）
2. 色指定のフォーマット制約が分かりづらい（HEX の扱い）
3. 文字色と背景色の概念が API 上で分離されていない
4. `draw_grid_border` の指定方法が直感とずれる（`range` 不可）
5. `out_path` のルート制約が分かりづらく、再試行がループする

## 目的

1. AI エージェントの初回成功率を上げる
2. 色指定を「任意HEXで安全に使える仕様」にする
3. `color`（文字色）と `fill_color`（背景色）を明確に分離する
4. 既存 API 互換を可能な限り維持しつつ入力許容を拡張する

## 非目的

1. `PatchOp` の大規模再設計
2. 既存成功ケースの挙動変更
3. 追加の外部依存導入

## スコープ

### FS-01: PatchOp 入力エイリアス正規化（非色情報）

`_coerce_patch_ops` で以下を正規化してから `PatchOp` 検証へ渡す。

1. `add_sheet`: `name` があり `sheet` が無い場合は `sheet = name`
2. `set_dimensions`: `col` -> `columns`, `row` -> `rows`
3. `set_dimensions`: `width` -> `column_width`（列指定時）, `height` -> `row_height`（行指定時）

優先順位:
1. 正式フィールドが存在する場合は正式フィールドを優先
2. エイリアスと正式フィールドが矛盾する場合はエラーにする

注意:
- `color -> fill_color` の自動変換は行わない（`color` と `fill_color` を別概念として扱う）

### FS-02: 色指定を任意HEXで許容（`color` / `fill_color` 共通）

`color` と `fill_color` は固定色ではなく、任意HEXを受け付ける。

許容フォーマット:
1. `RRGGBB`
2. `AARRGGBB`
3. `#RRGGBB`
4. `#AARRGGBB`

内部正規化:
1. `#` が無い場合は補完
2. 大文字化（例: `#1f4e79` -> `#1F4E79`）

不正文字列（桁数不一致・非16進文字）はエラー。

### FS-03: 文字色操作 `set_font_color` の追加

`PatchOpType` に `set_font_color` を追加する。

仕様:
1. 必須: `sheet`, `color`
2. 対象指定: `cell` または `range` のどちらか一方（両方不可）
3. `fill_color` は受け付けない
4. 既存のスタイル操作と同様に最大対象セル数制限を適用

意味:
1. `color`: 文字色（font color）
2. `fill_color`: 背景色（solid fill）

### FS-04: `draw_grid_border` の `range` shorthand 対応

`draw_grid_border` で `range` を受け取ったら内部で以下に変換する。

1. `base_cell`: 範囲左上セル
2. `row_count`: 範囲行数
3. `col_count`: 範囲列数

制約:
1. `range` と `base_cell/row_count/col_count` の併用は不可
2. 既存の最大セル数制限は維持

### FS-05: `out_path` の root 基準化と診断改善

1. 相対パス `out_path` は必ず MCP `--root` 基準で解決する
2. `Path is outside root` エラー時、以下を含むメッセージを返す
   - 解決後パス
   - 許可 root
   - 有効な指定例（相対）

例: `out_path: "outputs/book.xlsx"`

### FS-06: ランタイム情報取得ツール（任意）

新規ツール `exstruct_get_runtime_info` を追加し、以下を返す。

1. `root`
2. `cwd`
3. `platform`
4. `path_examples`（有効な相対/絶対例）

目的は path 制約デバッグの初動短縮。

## 受け入れ条件（Acceptance Criteria）

### AC-01 入力互換

1. `name` 指定の `add_sheet` が成功する
2. `col` + `width` 指定の `set_dimensions` が成功する

### AC-02 HEX自由指定

1. `fill_color: "1F4E79"` が成功し、内部値が `#1F4E79`
2. `color: "CC336699"` が成功し、内部値が `#CC336699`
3. 不正フォーマットは失敗し、エラーメッセージは明確

### AC-03 色概念の分離

1. `set_fill_color` は `fill_color` のみ受理し、文字色は変更しない
2. `set_font_color` は `color` のみ受理し、背景色は変更しない
3. `set_font_color` に `fill_color` を渡した場合はエラー

### AC-04 border shorthand

1. `draw_grid_border` で `range: "A4:G19"` が成功する
2. `range` と `base_cell` 併用時は明確なエラー

### AC-05 path UX

1. 相対 `out_path` が root 配下へ解決される
2. root 外パスで、修正に必要な情報を含むエラーが返る

### AC-06 後方互換

1. 既存の正式入力は挙動変更なし
2. 既存テストがすべて通る

## 影響範囲

1. `src/exstruct/mcp/server.py`（入力正規化）
2. `src/exstruct/mcp/patch_runner.py`（`set_font_color` 追加、色検証拡張）
3. `src/exstruct/mcp/io.py`（パス診断文言）
4. `src/exstruct/mcp/tools.py`（必要に応じたモデル説明更新）
5. `docs/mcp.md`（色指定と新opサンプル追加）
6. `tests/mcp/*`（回帰・新規テスト）

## リスクと対策

1. `color` の意味の誤解（文字色か背景色か）
   - 対策: `set_font_color` と `set_fill_color` の責務を明示し、混在入力はエラー
2. 入力曖昧性の増加
   - 対策: 正式フィールド優先、矛盾時エラー
3. 変換ロジック肥大化
   - 対策: 正規化関数を小分けし 1 責務を維持

## リリース方針（段階導入）

### Phase 1（必須）

1. FS-01
2. FS-02
3. FS-03
4. FS-04
5. FS-05

### Phase 2（推奨）

1. FS-06
2. ドキュメントの「失敗しやすい例」拡充
