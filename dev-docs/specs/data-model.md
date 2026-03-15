# ExStruct Data Model Specification

**Version**: 0.16
**Status**: Canonical

This document is the single canonical source for all models returned by ExStruct.
core / io / integrate must comply with this specification.
Models are implemented with **pydantic v2**.

---

# 1. Overview

ExStruct serializes Excel workbooks to JSON as **semantic structure** suitable for LLM consumption.
Unless otherwise noted, all models below are Pydantic `BaseModel` instances.

---

# 2. Shape / Arrow / SmartArt Models

The `shapes` output is a union of the three models below. Use `kind` to distinguish them.

```jsonc
BaseShape {
  id: int | null   // per-sheet sequential id (may be null for arrows)
  text: str
  l: int           // left (px)
  t: int           // top  (px)
  w: int | null    // width (px)
  h: int | null    // height (px)
  rotation: float | null
}

Shape extends BaseShape {
  kind: "shape"
  type: str | null // MSO shape type label
}

Arrow extends BaseShape {
  kind: "arrow"
  begin_arrow_style: int | null
  end_arrow_style: int | null
  begin_id: int | null // Shape.id of the connector start connection
  end_id: int | null   // Shape.id of the connector end connection
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

Notes:

- `direction` normalizes the direction of lines and arrows to 8 compass points
- Arrow styles correspond to Excel enums
- `begin_id` / `end_id` are the `id` of the shape the connector is connected to
- `SmartArtNode` is represented as a nested structure, with `nodes` as the tree root

---

# 3. CellRow Model

```jsonc
CellRow {
  r: int                                  // row number (1-based)
  c: { [colIndex: str]: str | int | float } // non-empty cells only; keys are column index strings
  links: { [colIndex: str]: url } | null    // only when hyperlink support is enabled
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

Series hold references rather than values, reducing payload size.

---

# 5. Chart Model

```jsonc
Chart {
  name: str
  chart_type: str              // label from XL_CHART_TYPE_MAP
  title: str | null
  y_axis_title: str
  y_axis_range: [float]        // [min, max], may be empty
  w: int | null
  h: int | null
  series: [ChartSeries]
  l: int                       // left (px)
  t: int                       // top  (px)
  error: str | null            // set only on parse failure
}
```

---

# 6. PrintArea Model

```jsonc
PrintArea {
  r1: int  // start row (1-based, inclusive)
  c1: int  // start column (0-based, inclusive)
  r2: int  // end row (1-based, inclusive)
  c2: int  // end column (0-based, inclusive)
}
```

Notes:

- Multiple areas may be held per sheet
- Included when obtainable in `standard` / `verbose`

---

# 7. PrintAreaView Model

```jsonc
PrintAreaView {
  book_name: str
  sheet_name: str
  area: PrintArea
  shapes: [Shape | Arrow | SmartArt]
  charts: [Chart]
  rows: [CellRow]          // only rows intersecting the range; empty columns dropped
  table_candidates: [str]  // table candidates fully contained within the range
}
```

Notes:

- Coordinates are sheet-relative by default. When `normalize` is specified, the range top-left is used as the origin.

---

# 8. MergedCells Model

```jsonc
MergedCells {
  schema: ["r1", "c1", "r2", "c2", "v"]
  items: [[int, int, int, int, str]]
}
```

- `items` is an array of `(r1, c1, r2, c2, v)` tuples
- Row is 1-based; column is 0-based
- `v` is the representative value of the merged cell; `" "` is output even when there is no value

---

# 9. SheetData Model

```jsonc
SheetData {
  rows: [CellRow]
  shapes: [Shape | Arrow | SmartArt]
  charts: [Chart]
  table_candidates: [str]
  print_areas: [PrintArea]
  auto_print_areas: [PrintArea] // auto page-break rectangles (COM required, disabled by default)
  formulas_map: {[formula: str]: [[int, int]]} // (row=1-based, col=0-based)
  colors_map: {[colorHex: str]: [[int, int]]} // (row=1-based, col=0-based)
  merged_cells: MergedCells | null
}
```

Notes:

- `table_candidates` are table detection results
- `print_areas` are defined print ranges
- `auto_print_areas` are obtained from Excel COM auto page breaks
- Merged cell value output in `rows` is controlled by the `include_merged_values_in_rows` flag (default: `True`)

---

# 10. WorkbookData Model (Top Level)

```jsonc
WorkbookData {
  book_name: str
  sheets: { [sheetName: str]: SheetData }
}
```

Notes:

- Sheet names are preserved verbatim as Excel Unicode names

---

# 11. Export Helpers (`SheetData` / `WorkbookData`)

Common:

- `to_json(pretty=False, indent=None, include_backend_metadata=False)`
- `to_yaml(include_backend_metadata=False)` (`pyyaml` required)
- `to_toon(include_backend_metadata=False)` (`python-toon` required)
- `save(path, pretty=False, indent=None, include_backend_metadata=False)`
  - Auto-detects format by extension: `.json` / `.yaml` / `.yml` / `.toon`
  - Unsupported extensions raise `ValueError`
- After `model_dump(exclude_none=True)`, remove empty values with `dict_without_empty_values`
- By default, backend metadata (`provenance`, `approximation_level`, `confidence`) is not included in serialized output

`SheetData`:

- `book_name` is not included when serialized (single sheet)

`WorkbookData`:

- Payload includes `book_name` and `sheets`
- `__getitem__(sheet_name)` retrieves a SheetData
- `__iter__()` yields `(sheet_name, SheetData)` in order

---

# 12. Versioning Principles

- Update this file before making model changes
- Keep models as pure data containers; do not add side effects
- core / io / integrate must only return models faithful to this specification and must not add custom fields

---

# 13. Changelog

- 0.3: Added serialize/save helpers; defined `__iter__` / `__getitem__` on `WorkbookData`
- 0.4: Added `CellRow.links` (hyperlinks are opt-in)
- 0.5: Added `PrintArea`; held in `SheetData.print_areas`
- 0.6: Extract PrintArea by default; table detection as before
- 0.7: Added size `w` / `h` to Chart
- 0.8: Added `SheetData.auto_print_areas` (COM auto page-break rectangles; disabled by default)
- 0.9: Added `name` / `begin_connected_shape` / `end_connected_shape` to Shape; later renamed to `begin_id` / `end_id`
- 0.10: Added `id` to Shape; removed `name`
- 0.11: Unified connector field names to `begin_id` / `end_id`
- 0.12: Added `SheetData.colors_map`
- 0.13: Split Shape into `Shape` / `Arrow` / `SmartArt`; added nested `SmartArtNode` structure
- 0.14: Added `MergedCell` / `SheetData.merged_cells`
- 0.15: Changed `MergedCells` to schema + items format introducing a compressed representation
- 0.16: Added `SheetData.formulas_map`

---

# Appendix A. MCP Patch Models

The model group used by MCP patch/make operations remains importable from
`exstruct.mcp.patch_runner` for backward compatibility.

The actual locations are as follows.

- Canonical models: `src/exstruct/mcp/patch/models.py`
- Compatibility facade: `src/exstruct/mcp/patch_runner.py`
- Service layer: `src/exstruct/mcp/patch/service.py`

Primary models:

- `PatchOp`
- `PatchRequest`
- `MakeRequest`
- `PatchResult`
- `PatchDiffItem`
- `PatchErrorDetail`
- `FormulaIssue`
