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

### mcp/patch（Patch 実装）

Patch 系は `src/exstruct/mcp/patch/` に責務分離して実装する。

- `patch_runner.py` → 互換性維持用ファサード（既存 import 経路を維持）
- `patch/legacy_runner.py` → 既存 patch 実装の後方互換レイヤ
- `patch/service.py` → `run_patch` / `run_make` のオーケストレーション
- `patch/runtime.py` → path/backend 選択など実行時ユーティリティ集約
- `patch/engine/openpyxl_engine.py` → openpyxl 実行境界
- `patch/engine/xlwings_engine.py` → xlwings(COM) 実行境界
- `patch/ops/openpyxl_ops.py` → openpyxl 向け op 適用入口
- `patch/ops/xlwings_ops.py` → xlwings 向け op 適用入口
- `patch/normalize.py` / `patch/specs.py` → op 正規化と仕様メタデータ
- `shared/a1.py` / `shared/output_path.py` → A1 と出力 path の共通処理

---

## AI エージェント向けガイド

- モデル変更は core の RawData 変換にも反映する
- 外部依存（openpyxl/xlwings）は core の境界で完結させる
- `PipelinePlan` は不変、実行状態は `PipelineState` へ分離
- 仕様は docs/DATA_MODEL.md を最優先とする
