# Excel 抽出仕様 - ExStruct

この文書は Excel からの抽出処理の最新仕様をまとめたものです。

## 全体フロー

1. `resolve_extraction_inputs` で include_* と mode を正規化
2. pre-com（openpyxl）で cells/print_areas/formulas_map/colors_map/merged_cells を取得
3. `standard` / `verbose` は com（xlwings）で shapes/charts/auto_page_breaks を取得
4. `libreoffice` は LibreOffice backend で shapes/charts を best-effort 抽出
5. COM 成功時は colors_map を COM 結果で上書き
6. rich backend 失敗時は cells+table_candidates を維持し、pre-com 取得済みの print_areas / formulas_map / colors_map / merged_cells もフラグに応じて保持する

## 座標系

- 行は 1-based
- 列は 0-based

## モード

- light: COM を完全にスキップし、cells+table_candidates を基本に pre-com artifact をフラグに応じて返す
- libreoffice: LibreOffice backend で rich artifact を best-effort 抽出し、失敗時は pre-com artifact を保持した cells fallback に戻る
- standard: 既存挙動（テキスト付き図形、必要に応じてチャート）
- verbose: 全図形 + サイズ付き、チャートもサイズ付き

## セル抽出

- pandas の `read_excel(header=None, dtype=str)` で読み込む
- 空白セルは無視
- 行データは `CellRow` に正規化

## テーブル抽出

- openpyxl のテーブル定義 + 罫線クラスターを統合
- COM が使えない場合でも table_candidates を維持

## Shapes / Arrows / SmartArt 抽出

抽出内容:

- Type / AutoShapeType の正規化（`type` は Shape のみ）
- Left/Top/Width/Height
- TextFrame2.TextRange.Text
- 矢印方向や接続情報
- SmartArt の layout/nodes/kids（ネスト構造）

## チャート抽出

抽出内容:

- ChartType（整数 → XL_CHART_TYPE_MAP で文字列化）
- Series / Axis Title / Axis Range
- Chart Title

## 印刷範囲 / 自動改ページ

- print_areas は pre-com（openpyxl）で取得し COM では補完のみ
- auto_page_breaks は COM のみで取得

## Colors Map

- 条件付き書式の色を含めるため COM を優先
- COM 成功時は COM 結果で上書き
- COM 失敗時のみ openpyxl 結果を使用

## エラーハンドリング / フォールバック

- COM / LibreOffice 不可・例外時は cells+table_candidates を返し、pre-com 取得済みの print_areas / formulas_map / colors_map / merged_cells は保持する
- fallback 理由は `FallbackReason` で統一ログ
