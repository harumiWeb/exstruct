# Excel Extraction Specification - ExStruct

この文書は Excel からの抽出処理の最新仕様をまとめたものです。

## 全体フロー

1. `resolve_extraction_inputs` で include_* と mode を正規化
2. pre-com（openpyxl）で cells/print_areas/colors_map を取得
3. com（xlwings）で shapes/charts/auto_page_breaks を取得
4. COM 成功時は colors_map を COM 結果で上書き
5. COM 失敗時は cells+tables のみでフォールバック

## 座標系

- 行は 1-based
- 列は 0-based

## モード

- light: COM を完全にスキップ、cells+tables のみ
- standard: 既存挙動（テキスト付き図形、必要に応じてチャート）
- verbose: 全図形 + サイズ付き、チャートもサイズ付き

## Cells

- pandas の `read_excel(header=None, dtype=str)` で読み込む
- 空白セルは無視
- 行データは `CellRow` に正規化

## Tables

- openpyxl のテーブル定義 + 罫線クラスターを統合
- COM が使えない場合でも table_candidates を維持

## Shapes

抽出内容:

- Type / AutoShapeType の正規化
- Left/Top/Width/Height
- TextFrame2.TextRange.Text
- 矢印方向や接続情報

## Charts

抽出内容:

- ChartType（整数 → XL_CHART_TYPE_MAP で文字列化）
- Series / Axis Title / Axis Range
- Chart Title

## Print Areas / Auto Page Breaks

- print_areas は pre-com（openpyxl）で取得し COM では補完のみ
- auto_page_breaks は COM のみで取得

## Colors Map

- 条件付き書式の色を含めるため COM を優先
- COM 成功時は COM 結果で上書き
- COM 失敗時のみ openpyxl 結果を使用

## エラーハンドリング / フォールバック

- COM 不可・例外時は cells+tables のみ返す
- fallback 理由は `FallbackReason` で統一ログ
