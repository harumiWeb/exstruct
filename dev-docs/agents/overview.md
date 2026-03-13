# ExStruct - Excel 構造抽出エンジン概要

**ExStruct** は Excel ブックから意味構造を抽出する Python ライブラリです。
openpyxl、Excel COM（xlwings）、LibreOffice backend を組み合わせ、LLM が扱いやすい構造データを生成します。

## 特徴

- パイプライン設計で抽出フローを統一
- mode（light/libreoffice/standard/verbose）で抽出粒度を切替
- openpyxl/COM/LibreOffice の backend 抽象化
- JSON/YAML/TOON 出力（依存がある場合）
- print_areas / auto_page_breaks の出力対応
- フォールバック理由を統一ログで可視化

## 抽出対象

- Cells（値/リンク/座標）
- Tables（候補範囲）
- Shapes / Arrows / SmartArt（位置/テキスト/矢印/レイアウト）
- Charts（Series/Axis/Type/Title）
- Print Areas / Auto Page Breaks
- Colors Map（条件付き書式を含む）

## 利用例（概要）

- `extract(path, mode="standard")` で WorkbookData を取得
- `process_excel` でファイル出力やディレクトリ出力
- CLI で `exstruct file.xlsx --format json` を利用

## ディレクトリ構成（概要）

```txt
docs/                 公開ドキュメント
dev-docs/             内部ドキュメント
src/exstruct/
  core/               抽出パイプラインと backend
  models/             Pydantic モデル
  io/                 JSON/YAML/TOON 出力
  render/             PDF/PNG 出力
  cli/                CLI
tests/                テスト
```

AI エージェントは `docs/` を公開契約として読み、内部挙動は `dev-docs/specs/` と `dev-docs/adr/` を使って補完します。
