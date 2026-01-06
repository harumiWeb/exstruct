# ExStruct データモデル仕様

**Version**: 0.15
**Status**: Authoritative

本ドキュメントは ExStruct が返す全モデルの唯一の正準ソースです。
core / io / integrate はこの仕様に従うこと。モデルは **pydantic v2** で実装します。

---

# 1. Overview

ExStruct は Excel ワークブックを LLM が扱いやすい **意味構造** として JSON 化します。
特記がない限り、以下のモデルはすべて Pydantic の `BaseModel` です。

---

# 2. Shape / Arrow / SmartArt Model

出力の `shapes` は下記 3 モデルのユニオンです。`kind` で判別します。

```jsonc
BaseShape {
  id: int | null   // sheet 連番 id（矢印は null の場合あり）
  text: str
  l: int           // left (px)
  t: int           // top  (px)
  w: int | null    // width (px)
  h: int | null    // height (px)
  rotation: float | null
}

Shape extends BaseShape {
  kind: "shape"
  type: str | null // MSO 図形タイプラベル
}

Arrow extends BaseShape {
  kind: "arrow"
  begin_arrow_style: int | null
  end_arrow_style: int | null
  begin_id: int | null // コネクタ始点の接続 Shape.id
  end_id: int | null   // コネクタ終点の接続 Shape.id
  direction: "E"|"SE"|"S"|"SW"|"W"|"NW"|"N"|"NE" | null
}

SmartArtNode {
  text: str
  kids: [SmartArtNode]
}

SmartArt extends BaseShape {
  kind: "smartart"
  layout: str
  nodes: [SmartArtNode]
}
```

補足:

- `direction` は線や矢印の向きを 8 方位に正規化
- 矢印スタイルは Excel の enum に対応
- `begin_id` / `end_id` はコネクタが接続している図形の `id`
- `SmartArtNode` はネスト構造で表現し、`nodes` がツリーの根

---

# 3. CellRow Model

```jsonc
CellRow {
  r: int                                  // 行番号 (1-based)
  c: { [colIndex: str]: str | int | float } // 非空セルのみ、キーは列インデックス文字列
  links: { [colIndex: str]: url } | null    // ハイパーリンク有効化時のみ
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
  r1: int  // 開始行 (1-based, inclusive)
  c1: int  // 開始列 (0-based, inclusive)
  r2: int  // 終了行 (1-based, inclusive)
  c2: int  // 終了列 (0-based, inclusive)
}
```

補足:

- シートごとに複数保持可能
- `standard` / `verbose` で取得できる場合に含まれる

---

# 7. PrintAreaView Model

```jsonc
PrintAreaView {
  book_name: str
  sheet_name: str
  area: PrintArea
  shapes: [Shape | Arrow | SmartArt]
  charts: [Chart]
  rows: [CellRow]          // 範囲に交差する行のみ、空列は落とす
  table_candidates: [str]  // 範囲内に収まるテーブル候補
}
```

補足:

- 座標はデフォルトでシート基準。`normalize` 指定時は範囲左上を原点に再基準化

---

# 8. MergedCells Model

```jsonc
MergedCells {
  schema: ["r1", "c1", "r2", "c2", "v"]
  items: [[int, int, int, int, str]]
}
```

- `items` は `(r1, c1, r2, c2, v)` の配列
- row は 1-based、col は 0-based
- `v` は結合セルの代表値。値がない場合でも `" "` を出力する

---

# 9. SheetData Model

```jsonc
SheetData {
  rows: [CellRow]
  shapes: [Shape | Arrow | SmartArt]
  charts: [Chart]
  table_candidates: [str]
  print_areas: [PrintArea]
  auto_print_areas: [PrintArea] // 自動改ページ矩形 (COM 前提、デフォルト無効)
  colors_map: {[colorHex: str]: [[int, int]]} // (row=1-based, col=0-based)
  merged_cells: MergedCells | null
}
```

補足:

- `table_candidates` はテーブル検知結果
- `print_areas` は定義済み印刷範囲
- `auto_print_areas` は Excel COM の自動改ページから取得
- `rows` の結合セル値の出力は `include_merged_values_in_rows` フラグで制御（既定: `True`）

---

# 10. WorkbookData Model (トップレベル)

```jsonc
WorkbookData {
  book_name: str
  sheets: { [sheetName: str]: SheetData }
}
```

補足:

- シート名は Excel の Unicode 名をそのまま保持

---

# 11. Export Helpers (SheetData / WorkbookData)

共通:

- `to_json(pretty=False, indent=None)`
- `to_yaml()` (`pyyaml` 必須)
- `to_toon()` (`python-toon` 必須)
- `save(path, pretty=False, indent=None)`
  - 拡張子 `.json` / `.yaml` / `.yml` / `.toon` を自動判別
  - 非対応拡張子は `ValueError`
- `model_dump(exclude_none=True)` 後に `dict_without_empty_values` で空値を除去

`SheetData`:

- シリアライズ時に `book_name` は含まない（シート単体）

`WorkbookData`:

- ペイロードに `book_name` と `sheets` を含む
- `__getitem__(sheet_name)` で SheetData を取得
- `__iter__()` で `(sheet_name, SheetData)` を順に返す

---

# 12. Versioning Principles

- モデル変更時は本ファイルを先に更新する
- モデルは純粋なデータコンテナとし、副作用を持たせない
- core / io / integrate は本仕様に忠実なモデルのみを返し、独自フィールドを追加しない

---

# 13. Changelog

- 0.3: serialize/save ヘルパー追加、`WorkbookData` に `__iter__` / `__getitem__` を定義
- 0.4: `CellRow.links` を追加（ハイパーリンクは opt-in）
- 0.5: `PrintArea` を追加し、`SheetData.print_areas` で保持
- 0.6: PrintArea をデフォルト抽出。テーブル検知は従来通り
- 0.7: Chart にサイズ `w` / `h` を追加
- 0.8: `SheetData.auto_print_areas` を追加（COM 自動改ページ矩形、デフォルト無効）
- 0.9: Shape に `name` / `begin_connected_shape` / `end_connected_shape` を追加し、後に `begin_id` / `end_id` に変更
- 0.10: Shape に `id` を追加し、`name` を削除
- 0.11: コネクタのフィールド名を `begin_id` / `end_id` に統一
- 0.12: `SheetData.colors_map` を追加
- 0.13: Shape を `Shape` / `Arrow` / `SmartArt` に分割し、`SmartArtNode` のネスト構造を追加
- 0.14: `MergedCell` / `SheetData.merged_cells` を追加
- 0.15: `MergedCells` を schema + items 形式に変更し圧縮形式を導入
