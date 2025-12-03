# Excel Extraction Specification — ExStruct

この文書は Excel から情報を抽出する処理の「厳密仕様」です。

## Cells

- pandas の `read_excel(header=None, dtype=str)` で読み込む
- 空白セルは無視
- 行データは辞書形式で保持
- シートごとに `List[CellRow]`

## Shapes

抽出内容：

- `Type` → shape_type（MSO_SHAPE_TYPE_MAP）
- `AutoShapeType`
- `Left/Top/Width/Height`
- `TextFrame2.TextRange.Text`

## Charts

抽出内容：

- ChartType（整数 → XL_CHART_TYPE_MAP で文字列化）
- Series.values（配列）
- Series.xvalues（散布図のみ）
- Axis Title / Range
- Chart Title

## エラーハンドリング

- グラフ解析失敗時は `error` に理由を入れる
- 図形にテキストがない場合は空文字
