# Architecture — ExStruct

ExStruct は複数のモジュールに分割し、責務を明確にしています。

## 全体構造

```txt
exstruct/
  core/
    cells.py
    shapes.py
    charts.py
    integrate.py
  models/
    __init__.py
    maps.py
  io/
  render/
  cli/
    main.py
```

## モジュール責務

### core/

Excel API（xlwings）から情報を抽出する層  
→ **外部依存が集まる場所**

- `cells.py` → pandas 読み込み、セル行の構造化
- `shapes.py` → ShapeType, AutoShapeType, 座標、テキスト抽出
- `charts.py` → Series, Axis, ChartType 抽出
- `integrate.py` → 座標を用いたセルへの図形の意味付与

### models/

Pydantic による「中間データ構造」  
→ **LLM が扱える最適な構造**

### io/

データの出力  
→ JSON / TOON / YAML

### render/

RAG 向けレンダリング  
→ PDF/PNG 出力

### cli/

CLI エントリポイント  
→ `exstruct extract file.xlsx --format json`

---

## AI エージェント向けガイド

- モデル定義を変更したら core 層も変更する必要がある
- core 層は xlwings と密結合
- models 層は絶対に “副作用なし”
- 関数は idempotent を保つこと
- 仕様は docs/DATA_MODEL.md を最優先とする
