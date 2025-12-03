# ExStruct — Excel Semantic Structure Extraction Engine

**ExStruct** は、Excel ブックから「意味を持った構造データ」を抽出するための  
Python 製 OSS ライブラリです。

通常の openpyxl や pandas では取得できない Excel の以下の情報を高精度に抽出します：

- **Cells**（テキスト、座標、意味ブロック）
- **Shapes**（AutoShapeType、座標、テキスト）
- **Charts**（Series, Axis, ChartType, Title）
- **Layout**（位置・サイズから意味的配置を推論）
- **Grid Structure**（今後追加：罫線・マージセル・表推定）

ExStruct の目的は、
**Excel の視覚構造（Visual Structure）を LLM が理解できるデータに変換すること**です。

## 特徴

- Excel Shapes の **種類・座標・AutoShapeType** まで完全抽出
- Excel Chart の **Series / Axis / ChartType / Title** を抽出
- セルと図形の **意味的対応付け（位置による関係推定）**
- PDF / PNG レンダリング機能（RAG 用途 /）
- 結果を **JSON** に出力可能 (将来的にyamlとtoonにも対応予定)
- RAG / AI エージェントで利用しやすいデータ構造

## 想定用途

- Excel マニュアルの構造化（RAG 取り込み）
- Excel 上のフローチャート意味解析
- グラフのデータ抽出と意味理解
- 企業内 DX ツールのデータ変換エンジン
- Excel 上の図・図形レイアウト解析

## ディレクトリ構成（推奨）

```bash
root: .
├── docs/
│   └── *.md # 仕様書
├── src/
│   └── exstruct/
│       ├── models/ # Pydanticモデル / shape type 辞書
│       ├── core/ # 抽出処理
│       ├── io/ # ファイル出力
│       ├── render/ # PDF/PNG出力
│       ├── __init__.py # ライブラリエントリポイント
│       ├── cli/ # CLIエントリポイント
│       └── py.typed
├── .python-version
├── pyproject.toml
├── README.md
└── uv.lock
```

AI エージェントは docs 以下の仕様を参照してコード生成します。
