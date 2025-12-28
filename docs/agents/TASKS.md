# Task List

未完了タスクは [ ]、完了タスクは [x]

## 目的
テストカバレッジ向上のため、優先施策に基づき段階的にテストを追加する。

## フェーズ0: 準備
- [x] 既存テスト構成・命名規則の確認（tests を階層化し日本語命名方針）
- [x] 既存のテストユーティリティ/fixture の再利用可否を整理（conftest: COM/Render 判定）
- [x] 追加対象の関数・分岐を一覧化（coverage.xml の低カバレッジ）

## フェーズ1: 純粋関数/ユーティリティ（高優先）
- [x] core/shapes.py: _should_include_shape の分岐テスト（standard/light/verbose）
- [x] core/shapes.py: compute_line_angle_deg, angle_to_compass, has_arrow の境界値テスト
- [ ] core/cells.py: _coerce_numeric_preserve_format の入力パターンテスト
- [ ] core/backends/com_backend.py: _split_csv_respecting_quotes の引用符対応テスト
- [ ] core/backends/com_backend.py: _normalize_area_for_sheet の補正ロジックテスト

## フェーズ2: パース/分岐ロジック（高優先）
- [ ] core/charts.py: parse_series_formula の正常系パターン（カンマ/セミコロン区切り）
- [ ] core/charts.py: parse_series_formula の異常系（空/非SERIES）
- [ ] core/pipeline.py: COM 不可時のフォールバック分岐（openpyxl 経由の結果検証）
- [ ] core/pipeline.py: COM 失敗時のフォールバック分岐（例外誘発の分岐確認）
- [ ] core/pipeline.py: include_auto_page_breaks フラグの計画分岐

## フェーズ3: バックエンド/検出ロジック（中優先）
- [ ] core/cells.py: detect_tables の .xlsx / .xls 分岐（openpyxl/COM の呼び分け）
- [ ] core/backends/openpyxl_backend.py: extract_print_areas 複数範囲の処理
- [ ] core/backends/openpyxl_backend.py: extract_colors_map の例外時 None 返却
- [ ] core/backends/com_backend.py: extract_print_areas の例外処理分岐

## フェーズ4: 疑似 COM オブジェクトでの網羅（中優先）
- [ ] core/shapes.py: get_shapes_with_position のフィルタリング動作（ダミー shape）
- [ ] core/charts.py: get_charts の Chart/Series 生成（ダミー chart）
- [ ] core/backends/com_backend.py: extract_auto_page_breaks の失敗分岐

## フェーズ5: 統合テスト（低優先）
- [ ] 既存の openpyxl 作成テストの延長（簡易 workbook の抽出/エクスポート）
- [ ] CLI/Engine の最小 E2E（JSON 生成の基本構造のみ検証）

## 進行ルール
- [ ] 追加テストは 1 テスト = 1 挙動を原則に小さく分割
- [ ] 例外系は「どの分岐を通したか」が明確になる命名にする
