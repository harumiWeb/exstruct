# Task List

未完了タスクは [ ]、完了タスクは [x]

## 目的

テストカバレッジ向上のため、優先施策に基づき段階的にテストを追加する。

## フェーズ 0: 準備

- [x] 既存テスト構成・命名規則の確認（tests を階層化し日本語命名方針）
- [x] 既存のテストユーティリティ/fixture の再利用可否を整理（conftest: COM/Render 判定）
- [x] 追加対象の関数・分岐を一覧化（coverage.xml の低カバレッジ）

## フェーズ 1: 純粋関数/ユーティリティ（高優先）

- [x] core/shapes.py: \_should_include_shape の分岐テスト（standard/light/verbose）
- [x] core/shapes.py: compute_line_angle_deg, angle_to_compass, has_arrow の境界値テスト
- [x] core/cells.py: \_coerce_numeric_preserve_format の入力パターンテスト
- [x] core/backends/com_backend.py: \_split_csv_respecting_quotes の引用符対応テスト
- [x] core/backends/com_backend.py: \_normalize_area_for_sheet の補正ロジックテスト

## フェーズ 2: パース/分岐ロジック（高優先）

- [x] core/charts.py: parse_series_formula の正常系パターン（カンマ/セミコロン区切り）
- [x] core/charts.py: parse_series_formula の異常系（空/非 SERIES）
- [x] core/pipeline.py: COM 不可時のフォールバック分岐（openpyxl 経由の結果検証）
- [x] core/pipeline.py: COM 失敗時のフォールバック分岐（例外誘発の分岐確認）
- [x] core/pipeline.py: include_auto_page_breaks フラグの計画分岐

## フェーズ 3: バックエンド/検出ロジック（中優先）

- [x] core/cells.py: detect_tables の .xlsx / .xls 分岐（openpyxl/COM の呼び分け）
- [x] core/backends/openpyxl_backend.py: extract_print_areas 複数範囲の処理
- [x] core/backends/openpyxl_backend.py: extract_colors_map の例外時 None 返却
- [x] core/backends/com_backend.py: extract_print_areas の例外処理分岐

## フェーズ 4: 疑似 COM オブジェクトでの網羅（中優先）

- [x] core/shapes.py: get_shapes_with_position のフィルタリング動作（ダミー shape）
- [x] core/charts.py: get_charts の Chart/Series 生成（ダミー chart）
- [x] core/backends/com_backend.py: extract_auto_page_breaks の失敗分岐

## フェーズ 5: 統合テスト（低優先）

- [x] 既存の openpyxl 作成テストの延長（簡易 workbook の抽出/エクスポート）
- [x] CLI/Engine の最小 E2E（JSON 生成の基本構造のみ検証）

## 追加カバレッジ強化: cells.py / workbook.py

### cells.py（低カバレッジ重点）
- [x] _normalize_color_key / _normalize_rgb の正規化パターン（#/ARGB/auto/theme/indexed）
- [x] _normalize_ignore_colors の空/重複/不正キーの除外
- [ ] _resolve_cell_background / _resolve_fill_color_key の既定背景/None/パターン種別分岐
- [ ] _color_to_key の theme/indexed/auto 分岐（openpyxl Color 風ダミー）
- [ ] _header_like_row / _table_signal_score / _count_nonempty_cells の境界ケース
- [ ] _collect_table_candidates_from_values の検出閾値と空入力（min_nonempty_cells）

### workbook.py（低カバレッジ重点）
- [ ] openpyxl_workbook の close 失敗時の例外抑制（close が例外でも継続）
- [ ] xlwings_workbook の既存ブック検出経路（_find_open_workbook が返す場合）
- [ ] _find_open_workbook の fullname 例外・resolve 例外のハンドリング
- [ ] _find_open_workbook の全体例外時に None を返す経路
