# ExStruct テスト要件仕様書

Version: 0.2  
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
11. パフォーマンス／メモリ  

---

# 2. 機能要件

## 2.1 セル抽出
- [CEL-01] 空セルを除外し、非空セルのみ `c` に出力する  
- [CEL-02] 行番号 `r` は 0 始まり  
- [CEL-03] 列キーは文字列インデックス `"0"`, `"1"` …  
- [CEL-04] 改行・タブを含むセルも正しく読み取れる  
- [CEL-05] Unicode（日本語・絵文字・サロゲート）を保持  
- [CEL-06] pandas 読み込みで dtype=string を強制  
- [CEL-07] シート全体規模でも性能劣化しない  
- [CEL-08] `_coerce_numeric_preserve_format` が int/float/非数を正しく判定  
- [CEL-09] `detect_tables_openpyxl` が openpyxl Table を検出  
- [CEL-10] `CellRow.links` は mode=verbose か include_cell_links=True で出力  

## 2.2 図形抽出
- [SHP-01] AutoShape の type を正規化  
- [SHP-02] TextFrame を正しく取得  
- [SHP-03] サイズ `w`,`h` は取得できない場合のみ null  
- [SHP-04] グループ図形は展開方針を一貫させる  
- [SHP-05] 座標 `l`,`t` は整数で取得しズームの影響を受けない  
- [SHP-07] 回転角度が Excel と一致  
- [SHP-09] begin/end_arrow_style が Excel ENUM と一致  
- [SHP-10] direction を 8 方位に分類  
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
- [CH-01] ChartType は XL_CHART_TYPE_MAP で文字列化  
- [CH-02] タイトル取得（なければ null）  
- [CH-03] y_axis_title 取得（なければ空）  
- [CH-04] 軸 min/max は float  
- [CH-05] 未設定軸は空リスト  
- [CH-06] name_range を参照式で出力（例: `=Sheet1!$B$1`）  
- [CH-07] x_range を参照式で出力  
- [CH-08] y_range を参照式で出力  
- [CH-09] 主要チャート種別（散布・棒など）を解析  
- [CH-10] 失敗時は `error` にメッセージを残しチャートは維持  

## 2.5 レイアウト統合
- [LAY-01] Shape のテキストを属する行に紐づける  
- [LAY-02] 列方向の簡易紐づけ（未実装なら skip）  
- [LAY-03] 1 行に複数 shape がある場合も順序を保持  
- [LAY-04] 図形なしなら空リスト  

---

# 3. モデル検証要件
- [MOD-01] すべてのモデルが `BaseModel` 継承  
- [MOD-02] 型が DATA_MODEL.md と完全一致  
- [MOD-03] Optional は未指定で None  
- [MOD-04] 数値は int/float に正規化  
- [MOD-05] direction Literal で不正値は ValidationError  
- [MOD-06] rows/shapes/charts/tables はデフォルト空リスト  
- [MOD-07] WorkbookData は `__getitem__` と順序付き iteration を提供  

---

# 4. 出力要件（JSON/YAML/TOON）
- [EXP-01] None/空文字/空リスト/空 dict は `dict_without_empty_values` で除去  
- [EXP-02] JSON 出力は UTF-8  
- [EXP-03] YAML 出力は sort_keys=False  
- [EXP-04] TOON 出力が正しく生成される  
- [EXP-05] WorkbookData → JSON → WorkbookData の往復で破壊的変更なし  
- [EXP-06] `export_sheets` がシート単位でファイル出力  
- [EXP-07] `to_json` が pretty インデント対応  
- [EXP-08] `save(path)` は拡張子で判別し、未対応拡張子は ValueError  
- [EXP-09] `to_yaml` / `to_toon` は依存未導入時 MissingDependencyError、導入済みなら成功  
- [EXP-10] OutputOptions の include_* で対象フィールドを除外し空リストも出力しない  
- [EXP-11] `print_areas_dir` / `save_print_area_views` で印刷範囲ごとのファイル出力（範囲なしなら書き出さない）  
- [EXP-12] PrintAreaView は範囲内の行のみ保持し、範囲外セル/リンクを除外  
- [EXP-13] PrintAreaView の table_candidates は範囲に完全に収まるもののみ  
- [EXP-14] normalize=True で行・列インデックスを印刷範囲起点に再基準化  
- [EXP-15] include_print_areas=False なら print_areas_dir があっても出力しない  
- [EXP-16] PrintAreaView は範囲と交差する shapes/charts のみ含め、サイズ不明図形は点扱い  
- [EXP-17] Chart.w/h は verbose で出力、standard ではデフォルト非表示（include_chart_size で制御）  
- [EXP-18] Shape.w/h は include_shape_size で制御し、デフォルト True は verbose のみ  
- [EXP-19] auto_page_breaks_dir 指定時に include_auto_page_breaks=True で auto_print_areas を取得（COM 必須）  
- [EXP-20] export_auto_page_breaks は auto_print_areas が空なら例外、存在時のみ書き出し  
- [EXP-21] save_auto_page_break_views は auto_print_areas を Sheet1#auto#1 などユニークキーで保存  
- [EXP-22] serialize_workbook は未対応フォーマットで SerializationError  
- [EXP-23] export/process API は output_path/sheets_dir/print_areas_dir/auto_page_breaks_dir に str/Path を渡しても正しく出力できる  

---

# 5. CLI 要件
- [CLI-01] `exstruct extract file.xlsx` が成功する  
- [CLI-02] `--format json/yaml/toml` が機能する  
- [CLI-03] `--image` で PNG 出力  
- [CLI-04] `--pdf` で PDF 出力  
- [CLI-05] 無効パス入力時も安全終了（クラッシュしない）
- [CLI-06] エラーメッセージが stdout/stderr に出力される
- [CLI-07] `--print-areas-dir` で印刷範囲ファイルを出力し、include_print_areas=False ならスキップ
- [CLI-08] Windows の cp932 環境（例: PYTHONIOENCODING=cp932）でも stdout 出力が UTF-8 を維持し、Unicode 文字を含む JSON を正しくパイプできる

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
- [ERR-09] 書き込み失敗は OutputError を送出し元例外は __cause__ に保持  

---

# 7. 回帰要件
- [REG-01] 既存フィクスチャの JSON 構造が過去版と一致  
- [REG-02] モデルのキー削除・名称変更を破壊的変更として検知  
- [REG-03] 方向推定アルゴリズム変更を検知  
- [REG-04] ChartSeries 参照解析が過去結果と一致  

---

# 8. 非機能要件
- パフォーマンス／メモリ目標は別途定義時に追記  

---

# 9. モード出力要件
- [MODE-01] CLI `--mode` / API `extract(..., mode=)` は light/standard/verbose のみ（デフォルト standard）  
- [MODE-02] light: セル+テーブルのみ、shapes/charts 空、COM 不使用  
- [MODE-03] standard: 既存挙動（テキスト付き図形・矢印、COM 有効ならチャート）  
- [MODE-04] verbose: 全図形（サイズ付き）とチャート（サイズ付き、特定除外を除く）  
- [MODE-05] process_excel は PDF/画像オプション併用でも mode を伝搬  
- [MODE-06] standard で既存フィクスチャに回帰し不要図形が増えない  
- [MODE-07] 無効 mode は処理開始前にエラー  
- [INT-01] COM オープン失敗時はセル+table_candidates へのフォールバック  
- [IO-05] dict_without_empty_values で None/空リスト/空 dict を除去しネストを保持  
- [RENDER-01] Excel+COM+pypdfium2 の PDF/PNG スモークテスト（環境で ON/OFF）  
- [MODE-08] light でも openpyxl で print_areas 抽出、デフォルト出力は除外（auto 判定）  
