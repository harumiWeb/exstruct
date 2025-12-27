# Architecture - ExStruct

ExStruct はパイプライン中心の抽出アーキテクチャで、
openpyxl と Excel COM（xlwings）を役割分担させています。

## 全体構造

```txt
exstruct/
  core/
    pipeline.py
    integrate.py
    modeling.py
    workbook.py
    backends/
      base.py
      openpyxl_backend.py
      com_backend.py
    cells.py
    shapes.py
    charts.py
    ranges.py
    logging_utils.py
  models/
    __init__.py
    maps.py
  io/
    serialize.py
  render/
  cli/
    main.py
```

## パイプライン設計

- `resolve_extraction_inputs` が include_* と mode を正規化
- `PipelinePlan` は pre-com / com の静的ステップ構成のみ保持
- 実行状態は `PipelineState` / `PipelineResult` に分離
- `run_extraction_pipeline` が COM 可否判定とフォールバックを一元管理

## モジュール責務

### core/

抽出の中心層（外部依存を集約）

- `pipeline.py` → 抽出フロー、COM 判定、fallback、raw 生成
- `backends/*` → openpyxl/COM の抽象化
- `cells.py` → セル抽出、テーブル検出、colors_map
- `shapes.py` → 図形抽出、方向推定
- `charts.py` → チャート解析
- `ranges.py` → 範囲解析の共通関数
- `workbook.py` → openpyxl/xlwings の contextmanager
- `modeling.py` → RawData から WorkbookData/SheetData を生成
- `integrate.py` → pipeline 呼び出しに特化した薄い入口

### models/

Pydantic による公開データ構造
（外部 API は BaseModel を返す）

### io/

出力フォーマット（JSON / YAML / TOON）とファイル書き込み

### render/

PDF/PNG 出力（RAG 用途）

### cli/

CLI エントリポイント

---

## AI エージェント向けガイド

- モデル変更は core の RawData 変換にも反映する
- 外部依存（openpyxl/xlwings）は core の境界で完結させる
- `PipelinePlan` は不変、実行状態は `PipelineState` へ分離
- 仕様は docs/DATA_MODEL.md を最優先とする
