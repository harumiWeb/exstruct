# ExStruct データモデル仕様

**Version**: 0.10  
**Status**: Authoritative — 本ドキュメントは ExStruct が返す全モデルの唯一の正準ソースです。  
core / io / integrate は必ずこの仕様に従うこと。モデルは **pydantic v2** で実装します。

---

# 1. Overview

ExStruct は Excel ワークブックを LLM が扱いやすい **意味構造（Semantic Structure）** として JSON 化します。  
特記がない限り、以下のモデルはすべて Pydantic の `BaseModel` です。

---

# 2. Shape Model

```jsonc
Shape {
  id: int | null   // sheet 内での通番 id（線・矢印は null の場合あり）
  text: str
  l: int           // left (px)
  t: int           // top  (px)
  w: int | null    // width (px)
  h: int | null    // height(px)
  type: str | null // MSO 図形タイプのラベル
  rotation: float | null
  begin_arrow_style: int | null
  end_arrow_style: int | null
  begin_id: int | null // コネクタ始点の接続先 Shape.id
  end_id: int | null   // コネクタ終点の接続先 Shape.id
  direction: "E"|"SE"|"S"|"SW"|"W"|"NW"|"N"|"NE" | null
}
```

補足:
- `direction` は線や矢印の向きを 8 方位に正規化したもの。
- 矢印スタイルは Excel の enum に対応。
- `begin_id` / `end_id` は、コネクタが接続している図形の `id`（Excel の `ConnectorFormat.BeginConnectedShape` / `EndConnectedShape` に対応）。
- 線や矢印の Shape では `id` が null になる場合があります。

---

# 3. CellRow Model

```jsonc
CellRow {
  r: int                                // 行番号（Excel 由来、1-based）
  c: { [colIndex: str]: str | int | float } // 非空セルのみ、列インデックスは文字列 ("0","1",...)
  links: { [colIndex: str]: url } | null    // ハイパーリンク（有効化時のみ）
}
```

---

# 4. ChartSeries Model

```jsonc
ChartSeries {
  name: str
  name_range: str | null
  x_range: str | null
  y_range: str | null
}
```

シリーズは値ではなく参照を保持し、ペイロードを削減します。

---

# 5. Chart Model

```jsonc
Chart {
  name: str
  chart_type: str              // XL_CHART_TYPE_MAP のラベル
  title: str | null
  y_axis_title: str
  y_axis_range: [float]        // [min, max]、空の可能性あり
  w: int | null
  h: int | null
  series: [ChartSeries]
  l: int                       // left (px)
  t: int                       // top  (px)
  error: str | null            // 解析失敗時のみセット
}
```

---

# 6. PrintArea Model

```jsonc
PrintArea {
  r1: int  // 開始行 (0-based, inclusive)
  c1: int  // 開始列 (0-based, inclusive)
  r2: int  // 終了行 (0-based, inclusive)
  c2: int  // 終了列 (0-based, inclusive)
}
```

補足:
- シートごとに複数保持可能。
- `standard` / `verbose` で取得できる場合に含まれる。

---

# 7. PrintAreaView Model

```jsonc
PrintAreaView {
  book_name: str
  sheet_name: str
  area: PrintArea
  shapes: [Shape]
  charts: [Chart]
  rows: [CellRow]          // 範囲に交差する行のみ、空列は落とす
  table_candidates: [str]  // 範囲内に収まるテーブル候補
}
```

補足:
- 座標はデフォルトでシート基準。`normalize` 指定時は範囲左上を原点に再基準化。

---

# 8. SheetData Model

```jsonc
SheetData {
  rows: [CellRow]
  shapes: [Shape]
  charts: [Chart]
  table_candidates: [str]
  print_areas: [PrintArea]
  auto_print_areas: [PrintArea] // 自動改ページ矩形（COM 前提、デフォルト無効）
}
```

補足:
- `table_candidates` はテーブル検出結果。
- `print_areas` は定義済み印刷範囲。`auto_print_areas` は Excel COM の自動改ページから取得し、明示的に有効化した場合のみ含まれる。

---

# 9. WorkbookData Model (トップレベル)

```jsonc
WorkbookData {
  book_name: str
  sheets: { [sheetName: str]: SheetData }
}
```

補足:
- シート名は Excel の Unicode 名をそのまま保持。

---

# 10. Export Helpers (SheetData / WorkbookData)

共通:
- `to_json(pretty=False, indent=None)`
- `to_yaml()`（`pyyaml` 必須）
- `to_toon()`（`python-toon` 必須）
- `save(path, pretty=False, indent=None)` — 拡張子から `.json` / `.yaml` / `.yml` / `.toon` を自動判別。非対応拡張子は `ValueError`。
- `model_dump(exclude_none=True)` 後に `dict_without_empty_values` で空値を除去。

`SheetData`:
- シリアライズ時に `book_name` は含まない（シート単体）。

`WorkbookData`:
- ペイロードに `book_name` と `sheets` を含む。
- `__getitem__(sheet_name)` で SheetData を取得、`__iter__()` で `(sheet_name, SheetData)` を順序付きで返す。

---

# 11. Versioning Principles（エージェント向け）

- モデル変更時は必ず本ファイルを先に更新する。
- モデルは純粋なデータコンテナとし、副作用を持たせない。
- core / io / integrate は本仕様に忠実なモデルのみを返し、独自フィールドを追加しない。

---

# 12. Changelog

- 0.3: serialize/save ヘルパーを追加、`WorkbookData` に `__iter__` / `__getitem__` を定義。
- 0.4: `CellRow.links` を追加（ハイパーリンクは opt-in、verbose でデフォルト有効）。
- 0.5: `PrintArea` を追加し、`SheetData.print_areas` で保持。standard / verbose で出力。
- 0.6: PrintArea をデフォルト抽出。テーブル検出は従来通り。
- 0.7: Chart にサイズフィールド `w` / `h`（optional）を追加。
- 0.8: `SheetData.auto_print_areas` を追加（COM の自動改ページ矩形、デフォルト無効）。ヘルパーとデフォルト挙動を明確化。
- 0.9: Shape に `name` / `begin_connected_shape` / `end_connected_shape` を追加し、コネクタの接続元/接続先を表現（後に `begin_id` / `end_id` に名称変更）。
- 0.10: Shape に `id` を追加し、コネクタの接続元/接続先を `id` 参照に変更し、`name` をペイロードから除去。
- 0.11: コネクタのフィールド名を `begin_id` / `end_id` にリネーム。
