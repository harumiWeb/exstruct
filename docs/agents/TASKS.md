# Task List

未完了タスクは [ ]、完了タスクは [x]

## ロードマップ（重要度順）

1. 抽出パイプラインの分割設計と共通インターフェース化
2. 範囲解析・テーブル検出・出力整形の共通化
3. リソース管理とログ/エラー処理の統一

## タスクリスト

- [x] 抽出ステップの責務分離（cells/tables/shapes/charts/print_areas/colors/auto_page_breaks）
  - [x] `core/pipeline.py` を新設し、ステップ定義の型（Protocol / BaseModel）を設計
  - [x] cells/tables/shapes/charts/print_areas/colors/auto_page_breaks を個別関数へ移動
  - [x] ステップごとの入力/出力モデル（中間構造）を定義
  - [x] `core/integrate.py` から抽出処理の実体を分離

- [ ] パイプライン定義の集約（mode 切替を 1 箇所に統一）
  - [ ] `ExtractionMode` と include_* フラグの解決ロジックをパイプライン層へ集約
  - [ ] light/standard/verbose の差分を「ステップ構成表」で管理
  - [ ] 既存の `engine.ExStructEngine.extract` から新パイプラインを呼び出し

- [x] openpyxl/COM の backend 抽象化（get_* 系 API を統一）
  - [x] `core/backends/base.py`（抽象）を設計
  - [x] `core/backends/openpyxl.py` を実装（cells/print_areas/colors/tables）
  - [x] `core/backends/com.py` を実装（shapes/charts/auto_page_breaks/colors/tables fallback）
  - [x] backend 選択規則（mode / SKIP_COM_TESTS / 拡張子）を整理

- [x] テーブル検出の共通化（矩形抽出・評価・範囲文字列化の再利用）
  - [x] 矩形抽出（border cluster 検出）を共通関数へ移動
  - [x] 矩形評価（density/coverage/header/score）を共通化
  - [x] 出力の範囲文字列化を共通化
  - [x] openpyxl/COM からの入力差分のみを残す

- [x] 範囲解析ユーティリティの統合（print area 解析の重複排除）
  - [x] `core/ranges.py` に `_parse_range_zero_based` 系を集約
  - [x] `integrate.py` / `io/__init__.py` の重複実装を置換
  - [x] 既存の範囲文字列の仕様差分を整理（"Sheet!A1:B2" など）

- [ ] 出力整形ロジックの共通化（json/yaml/toon 分岐とファイル書込みの整理）
  - [ ] `io/serialize.py` に format 判定・文字列生成を集約
  - [ ] `save_print_area_views` / `save_auto_page_break_views` / `save_sheets` を共通 writer に集約
  - [ ] 拡張子・命名規則を統一しテストを追加

- [ ] workbook の open/close 管理を共通化（contextmanager 化）
  - [ ] `core/workbook.py` に openpyxl/COM の contextmanager を用意
  - [ ] 例外時の close/quit を統一
  - [ ] `integrate.py` / `cells.py` からの散在呼び出しを置換

- [ ] モデル化フェーズの分離（生データ収集とモデル生成の段階化）
  - [ ] 生データ収集の出力を中間モデルで統一
  - [ ] `SheetData` / `WorkbookData` の生成を builder に移動
  - [ ] `integrate_sheet_content` を分割して責務整理

- [ ] フォールバック/警告ログの統一（理由コードの整理）
  - [ ] `errors.py` にフォールバック理由コードを定義
  - [ ] ログ出力の文言を統一しテストを追加
  - [ ] openpyxl/COM 切替時の warning を整理
