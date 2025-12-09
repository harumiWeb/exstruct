# ExStruct Data Model Specification

**Version**: 0.6  
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
  begin_arrow_style: int | null            // 始点の矢印タイプ (Excel enum)
  end_arrow_style: int | null              // 終点の矢印タイプ (Excel enum)
  direction: "E"|"SE"|"S"|"SW"|"W"|"NW"|"N"|"NE" | null
                                           // 向き（角度を8方位に正規化したもの）
}
```

### Notes

- `l`, `t`, `w`, `h` は座標・サイズで Excel レイアウト分析に使用する。
- `direction` は ExStruct が線の向きを 8 方位に正規化した結果。
- 矢印図形では begin/end_arrow_style に矢印種類が入る。

---

# 3. CellRow Model

Excel の行データを **“意味行単位”** として保存する。

```jsonc
CellRow {
  r: int                // 行番号 (0-based index)
  c: { [colIndex: str]: str | int | float }
                       // 非空セルのみを保持する dict
  links: { [colIndex: str]: url } | null // セルのハイパーリンク（mode=verbose で標準出力、その他は opt-in）
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
  w: int | null                // width (px) verbose で取得
  h: int | null                // height(px) verbose で取得
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

# 6. PrintArea Model

Excel の印刷範囲（矩形）を表現するモデル。

```jsonc
PrintArea {
  r1: int    // 開始行 (0-based, inclusive)
  c1: int    // 開始列 (0-based, inclusive)
  r2: int    // 終了行 (0-based, inclusive)
  c2: int    // 終了列 (0-based, inclusive)
}
```

### Notes

- 複数印刷範囲がある場合、シートごとに複数要素を保持。
- ExStruct の export では standard / verbose モードで標準出力・書き出しを行う。

---

# 7. PrintAreaView Model

印刷範囲ごとのスライスを保持するモデル。

```jsonc
PrintAreaView {
  book_name: str
  sheet_name: str
  area: PrintArea                // 対応する印刷範囲
  shapes: [Shape]
  charts: [Chart]
  rows: [CellRow]                // 範囲内にクリップした行（空列は落とす）
  table_candidates: [str]        // 範囲内に収まるテーブル候補のみ
}
```

### Notes

- 座標はデフォルトでシートのグローバル座標を維持。normalize オプション指定時は範囲起点で再基準化する。
- 形状・チャートは現状含めない（今後拡張予定）。

---

# 8. SheetData Model

Excel シート全体の意味構造。

```jsonc
SheetData {
  rows:   [CellRow]
  shapes: [Shape]
  charts: [Chart]
  table_candidates: [str]
  print_areas: [PrintArea]
}
```

### Notes

- `table_candidates` は v0.3 以降の "Table Detection" 機能のプレースホルダー。
- `print_areas` はデフォルト空リスト。mode=standard 以上で取得・出力（light では空）。

---

# 9. WorkbookData Model（トップレベル）

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

# 10. Versioning Principles（AI エージェント向け）

- モデル構造を変更する場合は **必ずこのファイルを更新すること**。
- ここにない属性を追加することは **許可しない**。
- 副作用を含むロジックは **models 層に書いてはならない**。
- core 層はこのモデルを「返すだけ」に徹する（推論はしてもよい）。

---

# 11. Export Helpers (SheetData / WorkbookData)

## 共通

- 両モデルにシリアライズヘルパーを実装済み。空値は `dict_without_empty_values` で除去される。
- `to_json(pretty=False, indent=None)`  
  - `pretty=True` かつ `indent` 未指定時は indent=2。デフォルトはコンパクト JSON。
- `to_yaml()`（pyyaml 未導入時は RuntimeError）
- `to_toon()`（python-toon 未導入時は RuntimeError）
- `save(path, pretty=False, indent=None)`  
  - 拡張子で自動判別: `.json` / `.yaml` / `.yml` / `.toon`  
  - 未対応拡張子は `ValueError`

## SheetData

- ペイロードは `model_dump(exclude_none=True)` 後に空値除去した dict。
- Save/serialize しても book_name は含まれない（SheetData 単体の内容のみ）。

## WorkbookData

- ペイロードは `book_name` + `sheets` を含む。シリアライズは `serialize_workbook` と同一ロジック。
- Save は `export` と同じ挙動（フォーマット判定・pretty オプションなど）を持つ。
- `__getitem__(sheet_name)` で SheetData を直接取得（KeyError は従来どおり）
- `__iter__()` で `(sheet_name, SheetData)` のペアを順に返す（sheets の順序を維持）

---

# 12. Changelog

- 0.3: モデルに出力ヘルパー (`to_json`/`to_yaml`/`to_toon`/`save`) を追加し、フォーマット判定・依存チェック・pretty 仕様を明文化。`WorkbookData` の `__iter__` / `__getitem__` を定義。
- 0.4: CellRow に `links` を追加（列インデックス→URL）。mode=verbose でデフォルト出力、その他は opt-in。
- 0.5: PrintArea モデルを追加し、SheetData に `print_areas` を導入。印刷範囲は verbose モード時のみ出力。
- 0.6: PrintArea をデフォルト取得し、standard/verbose で出力。印刷範囲ごとのスライス。
- 0.7: 印刷範囲内にchartが存在するか計算するためにChartにw/h属性をoptionalで追加。

---

以上が、ExStruct 最新バージョンの正準データモデル仕様となります。
