# ExStruct テスト要件仕様書

Version: 0.3
Status: Required for Release

ExStruct の全機能について、正式なテスト要件をまとめたドキュメントです。AI エージェント／人間開発者が自動テスト・手動テストを設計するための基盤とします。

---

# 1. カバレッジカテゴリ

1. セル抽出
2. 図形抽出
3. 矢印・方向推定
4. チャート抽出
5. レイアウト統合
6. Pydantic 検証
7. 出力（JSON/YAML/TOON）
8. CLI
9. エラー処理・フェイルセーフ
10. 回帰
11. パフォーマンス・メモリ

---

# 2. 機能要件

## 2.1 セル抽出

- [CEL-01] 空セルを除外し、非空セルのみ `c` に出力する
- [CEL-02] 行番号 `r` は 0 始まり
- [CEL-03] 列キーは 0 始まりのインデックス "0", "1" …
- [CEL-04] 改行・タブを含むセルも正しく読み取れる
- [CEL-05] Unicode（日本語・絵文字・サロゲート）を保持
- [CEL-06] pandas 読み込みで dtype=string を強制
- [CEL-07] シート全体規模でも性能劣化しない
- [CEL-08] `_coerce_numeric_preserve_format` が int/float/非数を正しく判定
- [CEL-09] `detect_tables_openpyxl` が openpyxl Table を検出
- [CEL-10] `CellRow.links` は mode=verbose または include_cell_links=True で出力
- [CEL-11] detect_tables は .xlsx/.xls の拡張子と openpyxl 有無で経路を切り替える

## 2.1.1 セル背景色

- [COL-01] `include_colors_map=True` のみ `colors_map` を抽出
- [COL-02] `include_default_background=False` では `FFFFFF` を出力しない
- [COL-03] `ignore_colors` 指定時は対象色を除外（`#` 付き/大小文字を正規化）
- [COL-04] COM 利用時は `DisplayFormat.Interior` を参照し条件付き書式を含めて取得
- [COL-05] `_normalize_color_key` / `_normalize_rgb` は ARGB/#/auto/theme/indexed を正規化

## 2.2 図形抽出

- [SHP-01] AutoShape の type を正規化
- [SHP-02] TextFrame を正しく取得
- [SHP-03] サイズ `w`,`h` は取得できない場合のみ null
- [SHP-04] グループ図形は展開方針を一貫させる
- [SHP-05] 座標 `l`,`t` は整数で取得しズームの影響を受けない
- [SHP-07] 回転角度は Excel と一致
- [SHP-09] begin/end_arrow_style は Excel ENUM と一致
- [SHP-10] direction は 8 方位に正規化
- [SHP-11] テキストなし図形は text=""
- [SHP-12] 複数段落のテキストも取得

## 2.3 矢印方向推定

- [DIR-01] 0° ±22.5° → "E"
- [DIR-02] 45° ±22.5° → "NE"
- [DIR-03] 90° ±22.5° → "N"
- [DIR-04] 135° ±22.5° → "NW"
- [DIR-05] 180° ±22.5° → "W"
- [DIR-06] 225° ±22.5° → "SW"
- [DIR-07] 270° ±22.5° → "S"
- [DIR-08] 315° ±22.5° → "SE"
- [DIR-09] 境界角度は仕様どおり丸める

## 2.4 チャート抽出

- [CH-01] ChartType は `XL_CHART_TYPE_MAP` で正規化
- [CH-02] タイトル取得（なければ null）
- [CH-03] y_axis_title 取得（なければ空文字）
- [CH-04] 軸 min/max は float
- [CH-05] 未設定軸は空リスト
- [CH-06] name_range を参照式で出力（例: `=Sheet1!$B$1`）
- [CH-06a] 文字列リテラルの series 名は name_literal に格納される
- [CH-07] x_range を参照式で出力
- [CH-08] y_range を参照式で出力
- [CH-09] 主要チャート種別（散布・棒など）を解析
- [CH-10] 失敗時は `error` にメッセージを残しチャートを維持
- [CH-11] 文化圏差のセミコロン区切りも解析できる

## 2.5 レイアウト統合

- [LAY-01] Shape のテキストを属する行に紐づける
- [LAY-02] 列方向の簡易紐づけ（未実装のため skip）
- [LAY-03] 1 行に複数 shape がある場合も順序を保持
- [LAY-04] 図形なしなら空リスト

---

# 3. モデル検証要件

- [MOD-01] すべてのモデルは `BaseModel` 継承
- [MOD-02] 型が DATA_MODEL.md と完全一致
- [MOD-03] Optional は未指定で None
- [MOD-04] 数値は int/float に正規化
- [MOD-05] direction Literal で不正値は ValidationError
- [MOD-06] rows/shapes/charts/tables はデフォルト空リスト
- [MOD-07] WorkbookData は `__getitem__` と順序付き iteration を提供
- [MOD-08] PrintArea は row=1-based / column=0-based を満たす

---

# 4. 出力要件（JSON/YAML/TOON）

- [EXP-01] None/空文字/空リスト/空 dict は `dict_without_empty_values` で除去
- [EXP-02] JSON 出力は UTF-8
- [EXP-03] YAML 出力は sort_keys=False
- [EXP-04] TOON 出力が正しく生成される
- [EXP-05] WorkbookData → JSON → WorkbookData の往復で破壊的変更なし
- [EXP-06] `export_sheets` がシート単位でファイル出力
- [EXP-07] `to_json` は pretty/indent に対応
- [EXP-08] `save(path)` は拡張子で判別し未対応拡張子は ValueError
- [EXP-09] `to_yaml` / `to_toon` は依存未導入時に MissingDependencyError
- [EXP-10] OutputOptions の include\_\* で対象フィールドを除外し空リストは出力しない
- [EXP-11] `print_areas_dir` / `save_print_area_views` で印刷範囲ごとのファイル出力（範囲なしなら書き出さない）
- [EXP-12] PrintAreaView は範囲内の行のみ保持し、範囲外のセル/リンクを除外
- [EXP-13] PrintAreaView の table_candidates は範囲内に完全に収まるもののみ
- [EXP-14] normalize=True で行・列インデックスを印刷範囲起点に再基準化
- [EXP-15] include_print_areas=False の場合は print_areas_dir があっても出力しない
- [EXP-16] PrintAreaView は範囲と交差する shapes/charts のみ含め、サイズ不明の図形は点扱い
- [EXP-17] Chart.w/h は verbose で出力し、standard では include_chart_size で制御
- [EXP-18] Shape.w/h は include_shape_size で制御し、デフォルト True は verbose のみ
- [EXP-19] auto_page_breaks_dir 指定時は include_auto_page_breaks=True で auto_print_areas を取得（COM 必須）
- [EXP-20] export_auto_page_breaks は auto_print_areas が空なら例外、存在時のみ書き出し
- [EXP-21] save_auto_page_break_views は auto_print_areas を Sheet1#auto#1 などユニークキーで保存
- [EXP-22] serialize_workbook は未対応フォーマットで SerializationError
- [EXP-23] export/process API は output_path/sheets_dir/print_areas_dir/auto_page_breaks_dir に str/Path を渡しても正しく出力できる
- [EXP-24] fmt="yml" は yaml として扱い、拡張子は .yaml になる

---

# 5. CLI 要件

- [CLI-01] `exstruct extract file.xlsx` が成功する
- [CLI-02] `--format json/yaml/toon` が機能する
- [CLI-03] `--image` で PNG 出力
- [CLI-04] `--pdf` で PDF 出力
- [CLI-05] 無効パス入力時も安全に終了（クラッシュしない）
- [CLI-06] エラーメッセージは stdout/stderr に出力
- [CLI-07] `--print-areas-dir` で印刷範囲ファイルを出力し、include_print_areas=False ならスキップ
- [CLI-08] Windows の cp932 環境（例: PYTHONIOENCODING=cp932）でも stdout 出力が UTF-8 を維持

---

# 6. エラー処理要件

- [ERR-01] xlwings COM エラーでもプロセスが落ちない
- [ERR-02] 図形抽出失敗でも他要素を維持
- [ERR-03] チャート抽出失敗時は Chart.error に記録
- [ERR-04] 壊れた参照範囲は例外にせず null/error を記録
- [ERR-05] Excel ファイルを開けない場合にメッセージを出し終了
- [ERR-06] openpyxl `_print_area` 設定も抽出漏れしない
- [ERR-07] auto_print_areas が空の場合 export_auto_page_breaks は PrintAreaError（ValueError 互換）を送出
- [ERR-08] YAML/TOON 依存なしの場合 MissingDependencyError でインストール案内
- [ERR-09] 書き込み失敗は OutputError を送出し、例外の **cause** に保持

---

# 7. 回帰要件

- [REG-01] 既存フィクスチャの JSON 構造が過去版と一致
- [REG-02] モデルのキー削除・名称変更を破壊的変更として検知
- [REG-03] 方向推定アルゴリズム変更を検知
- [REG-04] ChartSeries 参照解析が過去結果と一致

---

# 8. 非機能要件

- パフォーマンス・メモリ目標は別途定義時に追記

---

# 9. モード/統合要件

- [MODE-01] CLI `--mode` / API `extract(..., mode=)` は light/standard/verbose のみ（デフォルト standard）
- [MODE-02] light: セル+テーブルのみ、shapes/charts 空、COM 不使用
- [MODE-03] standard: 既存挙動（テキスト付き図形・矢印、COM 有効ならチャート）
- [MODE-04] verbose: 全図形（サイズ付き）とチャート（サイズ付き、特定除外を除く）
- [MODE-05] process_excel は PDF/画像オプション併用で mode を伝搬
- [MODE-06] standard で既存フィクスチャに回帰し不要図形が増えない
- [MODE-07] 無効 mode は処理開始前にエラー
- [INT-01] COM オープン失敗時はセル+table_candidates へのフォールバック
- [INT-02] COM フォールバック時も print_areas を保持する
- [IO-05] dict_without_empty_values で None/空リスト/空 dict を除去しネストを保持
- [RENDER-01] Excel+COM+pypdfium2 の PDF/PNG スモークテスト（環境 ON/OFF）
- [MODE-08] light では openpyxl で print_areas 抽出、デフォルト出力は除外（auto 判定）

## 2.6 Pipeline

- [PIPE-01] build*pre_com_pipeline は include*\* と mode に応じて必要なステップのみ含む
- [PIPE-02] build_cells_tables_workbook は print_areas を条件に反映し table_candidates を保持
- [PIPE-03] resolve_extraction_inputs は mode デフォルトで include_* を解決する
- [PIPE-04] run_extraction_pipeline は COM を試行し、失敗時は cells+tables にフォールバックする
- [PIPE-05] colors_map は COM 成功時に COM 結果で上書きし、失敗時のみ openpyxl を使う
- [PIPE-06] print_areas は openpyxl の結果を保持し、COM は不足分のみ補完する
- [PIPE-07] PipelineState は com_attempted/com_succeeded/fallback_reason を保持する
- [PIPE-08] include_auto_page_breaks=False の場合は auto_page_breaks の COM ステップを含めない
- [PIPE-MOD-01] build_workbook_data は raw コンテナから WorkbookData/SheetData を構築する
- [PIPE-MOD-02] collect_sheet_raw_data は抽出済みデータを raw コンテナにまとめる

## 2.7 Backend

- [BE-01] OpenpyxlBackend は include_links の有無で cells 抽出経路を切り替える
- [BE-02] OpenpyxlBackend は table 検出失敗時に空リストで継続する
- [BE-03] ComBackend は colors_map 抽出失敗時に None を返す
- [BE-04] OpenpyxlBackend は colors_map 抽出失敗時に None を返す
- [BE-05] ComBackend は print_areas 抽出失敗時に空マップで継続する

## 2.8 Ranges

- [RNG-01] parse_range_zero_based は "Sheet1!A1:B2" のようなシート付き範囲を正規化できる

## 2.9 Table Detection

- [TBL-01] 矩形マージは包含関係の矩形を統合しない
- [TBL-02] 値行列からテーブル候補の範囲文字列を生成できる

## 2.10 Workbook

- [WB-01] openpyxl_workbook は例外の有無に関係なく close を呼び出す
- [WB-02] openpyxl_workbook は既知の openpyxl 警告を抑制するフィルタを設定する
- [WB-03] _find_open_workbook は fullname/resolve 例外を許容し None を返す
- [WB-04] _find_open_workbook の全体例外時は None を返す
- [WB-05] xlwings_workbook は既存ブックが見つかれば App を起動しない

## 2.11 Logging

- [LOG-01] log_fallback は理由コードを含む警告ログを出力する

## 2.12 統合/E2E

- [E2E-01] light 抽出→serialize_workbook→export_sheets の一連が成功する
- [E2E-02] Engine.process は output_path=None のとき stream へ JSON を出力できる

---

# 10. COM テスト運用（ローカル手動）

CI では Excel COM を実行できないため、COM テストはローカル手動で実行する。
Codecov では unit と com を分離し、com は carryforward で維持する。

## 10.1 ローカル実行手順

- unit（CI 相当）: `task test-unit`
- COM: `task test-com`

## 10.2 Codecov 手動アップロード（任意）

手動アップロード時は `CODECOV_TOKEN` と `CODECOV_SHA` を設定する。

- unit 送信: `codecov-cli upload-process -f coverage.xml -F unit -C %CODECOV_SHA% -t %CODECOV_TOKEN%`
- COM 送信: `codecov-cli upload-process -f coverage.xml -F com -C %CODECOV_SHA% -t %CODECOV_TOKEN%`
