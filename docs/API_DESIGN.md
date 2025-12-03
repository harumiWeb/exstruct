# API Design — ExStruct

ExStruct のトップレベル Python API 設計。

## 基本 API

```python
import exstruct as xs

data = xs.extract("file.xlsx")
```

### 戻り値

`WorkbookData`（Pydantic model）

---

## CLI

```bash
exstruct extract file.xlsx --format json
```

オプション：

- `--image` → PNG 出力
- `--pdf` → PDF 出力
- `--dpi` → PNG 解像度
- `--multiple` → 複数ファイル処理

---

## JSON/YAML/TOML 出力

```python
xs.export(data, "output.json")
```

---

## 今後拡張予定

- `xs.detect_tables()`
- `xs.semantic_shapes()`
- `xs.smartart()`
