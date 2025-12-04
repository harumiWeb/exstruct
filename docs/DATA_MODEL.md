# ExStruct Data Model Specification

**Version**: 0.2  
**Status**: Authoritative — このファイルは ExStruct 内の全てのモデル定義の唯一の正準ソースです。  
core・io・integrate モジュールは、必ずこの仕様に一致するように実装してください。

---

# 1. Overview

ExStruct のデータ構造は、Excel ファイルから抽出された **意味構造（Semantic Structure）** を  
LLM で扱いやすい JSON 形式で表現するためのものです。

このモデルは pydantic v2 ベースです。

---

# 2. Shape Model

Excel 上の図形（AutoShape, Line, Arrow, Group 等）の構造定義。

```jsonc
Shape {
  text: str                                // 図形内テキスト (空も可)
  l: int                                   // left (px)
  t: int                                   // top  (px)
  w: int | null                            // width (px)
  h: int | null                            // height(px)
  type: str | null                         // MSO Shape Type 名
  rotation: float | null                   // 回転角度 (0–360)
  angle_deg: float | null                  // 方向線/矢印の角度
  begin_arrow_style: int | null            // 始点の矢印タイプ (Excel enum)
  end_arrow_style: int | null              // 終点の矢印タイプ (Excel enum)
  direction: "E"|"SE"|"S"|"SW"|"W"|"NW"|"N"|"NE" | null
                                           // 向き（角度を8方位に正規化したもの）
}
```

### Notes

- `l`, `t`, `w`, `h` は座標・サイズで Excel レイアウト分析に使用する。
- `direction` は `angle_deg` をもとに ExStruct が自動分類した 8 方位。
- 矢印図形では begin/end_arrow_style に矢印種類が入る。

---

# 3. CellRow Model

Excel の行データを **“意味行単位”** として保存する。

```jsonc
CellRow {
  r: int                // 行番号 (0-based index)
  c: { [colIndex: str]: str | int | float }
                       // 非空セルのみを保持する dict
}
```

### Notes

- `c` のキーは列インデックスを文字列化したもの `"0"`, `"1"`, `"5"` など。
- 必要最小限のトークン量で構造保持する設計。

---

# 4. ChartSeries Model

Excel Chart Series の定義。  
値そのものは保持せず、**参照範囲**を保持する構造に刷新した仕様。

```jsonc
ChartSeries {
  name: str                    // series の名前
  name_range: str | null       // =Sheet1!$B$1 など
  x_range: str | null          // XValues の参照範囲
  y_range: str | null          // Values  の参照範囲
}
```

### Notes

- 値を直接保持しないことでデータ量削減。
- 再現性の高い構造的メタデータが優先される。

---

# 5. Chart Model

Excel Chart（棒、折れ線、散布図、円、等）の構造化モデル。

```jsonc
Chart {
  name: str
  chart_type: str              // XL_CHART_TYPE_MAP の文字列表現
  title: str | null
  y_axis_title: str
  y_axis_range: [float]        // [min, max] 省略可
  series: [ChartSeries]
  l: int                       // left (px)
  t: int                       // top  (px)
  error: str | null            // 解析失敗時のみ
}
```

### Notes

- Chart の位置情報 (l,t) はレイアウト解析で利用。
- `error` が null でない場合、その Chart は完全解析できていない。

---

# 6. SheetData Model

Excel シート全体の意味構造。

```jsonc
SheetData {
  rows:   [CellRow]
  shapes: [Shape]
  charts: [Chart]
  table_candidates: [str]
}
```

### Notes

- `table_candidates` は v0.3 以降の "Table Detection" 機能のプレースホルダー。
- 空要素は出力時に除外される（dict_without_empty_values）。

---

# 7. WorkbookData Model（トップレベル）

すべてのシートをまとめた構造。

```jsonc
WorkbookData {
  book_name: str
  sheets: {
    [sheetName: str]: SheetData
  }
}
```

### Notes

- `sheets` は順序保証される dict。
- シート名は Excel そのままの Unicode 名を保持。

---

# 8. Versioning Principles（AI エージェント向け）

- モデル構造を変更する場合は **必ずこのファイルを更新すること**。
- ここにない属性を追加することは **許可しない**。
- 副作用を含むロジックは **models 層に書いてはならない**。
- core 層はこのモデルを「返すだけ」に徹する（推論はしてもよい）。

---

# 9. Changelog

なし

---

以上が、ExStruct 最新バージョンの正準データモデル仕様となります。
