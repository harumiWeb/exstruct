# ExStruct Test Requirements Specification

Version: 0.2  
Status: Required for Release

縺薙・譁・嶌縺ｯ縲・xStruct 縺ｮ縺吶∋縺ｦ縺ｮ讖溯・縺ｫ蟇ｾ縺吶ｋ **豁｣蠑上↑繝・せ繝郁ｦ∽ｻｶ荳隕ｧ** 縺ｧ縺ゅｊ縲・ 
AI 繧ｨ繝ｼ繧ｸ繧ｧ繝ｳ繝医・莠ｺ髢馴幕逋ｺ閠・曙譁ｹ縺悟盾辣ｧ縺励※閾ｪ蜍・謇句虚繝・せ繝医ｒ逕滓・縺ｧ縺阪ｋ繧医≧縺ｫ險ｭ險医＆繧後※縺・∪縺吶・

---

# 1. Test Coverage Categories

ExStruct 縺ｮ繝・せ繝医・莉･荳九・繧ｫ繝・ざ繝ｪ縺ｫ蛻・｡槭＆繧後ｋ・・

1. **繧ｻ繝ｫ謚ｽ蜃ｺ・・ells Extraction・・*
2. **蝗ｳ蠖｢謚ｽ蜃ｺ・・hapes Extraction・・*
3. **遏｢蜊ｰ繝ｻ譁ｹ蜷第耳螳夲ｼ・rrow + Direction Detection・・*
4. **繝√Ε繝ｼ繝域歓蜃ｺ・・hart Extraction・・*
5. **諢丞袖莉倅ｸ趣ｼ・ayout Integration・・*
6. **繝・・繧ｿ繝｢繝・Ν貅匁侠繝・せ繝茨ｼ・ydantic Validation・・*
7. **蜃ｺ蜉帙ヵ繧ｩ繝ｼ繝槭ャ繝茨ｼ・SON/YAML/TOML Writer・・*
8. **CLI 繝・せ繝・*
9. **繧ｨ繝ｩ繝ｼ蜃ｦ逅・・繝輔ぉ繧､繝ｫ繧ｻ繝ｼ繝・*
10. **蝗槫ｸｰ繝・せ繝茨ｼ・egression・・*
11. **繝代ヵ繧ｩ繝ｼ繝槭Φ繧ｹ/繝｡繝｢繝ｪ隕∽ｻｶ**

---

# 2. Functional Test Requirements (隧ｳ邏ｰ隕∽ｻｶ)

---

## **2.1 Cells Extraction Requirements**

### 蠢・医ユ繧ｹ繝・

- [CEL-01] 遨ｺ繧ｻ繝ｫ繧帝勁螟悶＠縲・撼遨ｺ繧ｻ繝ｫ縺ｮ縺ｿ `c` 縺ｫ蜃ｺ蜉帙＆繧後ｋ
- [CEL-02] 陦檎分蜿ｷ `r` 縺・0-based index 縺ｧ豁｣縺励￥蜃ｺ蜉帙＆繧後ｋ
- [CEL-03] 蛻礼分蜿ｷ縺・`"0"`, `"1"` 縺ｮ **譁・ｭ怜・繧ｭ繝ｼ** 縺ｧ蜃ｺ蜉帙＆繧後ｋ
- [CEL-04] 繧ｻ繝ｫ縺ｫ謾ｹ陦後・繧ｿ繝悶′蜷ｫ縺ｾ繧後※繧よｭ｣縺励￥隱ｭ縺ｿ霎ｼ繧√ｋ
- [CEL-05] Unicode・育ｵｵ譁・ｭ励∵律譛ｬ隱槭∫焚菴灘ｭ暦ｼ峨そ繝ｫ縺ｮ隱ｭ縺ｿ蜿悶ｊ
- [CEL-06] Pandas 隱ｭ縺ｿ霎ｼ縺ｿ縺ｫ繧医ｋ dtype=string 蠑ｷ蛻ｶ縺悟ｮ医ｉ繧後※縺・ｋ
- [CEL-07] 繧ｻ繝ｫ遽・峇縺悟､ｧ縺阪＞繝輔ぃ繧､繝ｫ縺ｧ繧・1 荳・そ繝ｫ遞句ｺｦ縺ｧ諤ｧ閭ｽ蝠城｡後′縺ｪ縺・
- [CEL-08] `_coerce_numeric_preserve_format` 縺梧紛謨ｰ繝ｻ蟆乗焚繝ｻ髱樊焚蛟､繧呈ｭ｣縺励￥蛻､螳壹☆繧・
- [CEL-09] `detect_tables_openpyxl` 縺・openpyxl 縺ｮ Table 繧ｪ繝悶ず繧ｧ繧ｯ繝医ｒ讀懷・縺ｧ縺阪ｋ
- [CEL-10] CellRow.links 縺悟・繧､繝ｳ繝・ャ繧ｯ繧ｹ 竊旦RL 縺ｧ譬ｼ邏阪＆繧後［ode=verbose 縺ｾ縺溘・ include_cell_links=True 縺ｮ縺ｨ縺阪・縺ｿ蜃ｺ蜉帙＆繧後ｋ

---

## **2.2 Shapes Extraction Requirements**

### 蝓ｺ譛ｬ蠖｢迥ｶ

- [SHP-01] AutoShape 縺ｮ type 縺梧ｭ｣縺励￥譁・ｭ怜・蛹悶＆繧後ｋ
- [SHP-02] TextFrame 縺ｮ譁・ｭ怜・縺梧ｭ｣縺励￥隱ｭ縺ｿ蜿悶ｌ繧・
- [SHP-03] 繧ｵ繧､繧ｺ・・, h・峨′ null 縺ｫ縺ｪ繧峨↑縺・ｼ亥叙蠕嶺ｸ榊庄譎ゅ・ null・・
- [SHP-04] Group・医げ繝ｫ繝ｼ繝怜峙蠖｢・牙・縺ｮ蟄仙峙蠖｢縺後☆縺ｹ縺ｦ螻暮幕縺輔ｌ繧・or 辟｡隕匁婿驥昴′邯ｭ謖√＆繧後ｋ

### 蠎ｧ讓・

- [SHP-05] l, t・・eft, top・峨′謨ｴ謨ｰ縺ｧ蜿門ｾ励＆繧後ｋ
- [SHP-06] 陦ｨ遉ｺ蛟咲紫繧・え繧｣繝ｳ繝峨え繧ｺ繝ｼ繝縺悟､峨ｏ縺｣縺ｦ繧ょｺｧ讓吶′螟牙虚縺励↑縺・

### 蝗櫁ｻ｢ / 遏｢蜊ｰ

- [SHP-07] rotation 縺・Excel 縺ｮ蝗櫁ｻ｢隗貞ｺｦ蛟､縺ｫ荳閾ｴ縺吶ｋ
- [SHP-09] begin_arrow_style / end_arrow_style 縺・Excel 縺ｮ ENUM 縺ｨ荳閾ｴ縺吶ｋ
- [SHP-10] direction 縺・8 譁ｹ菴榊・鬘槭↓蠕薙＞豁｣縺励￥邂怜・縺輔ｌ繧・

### 繝・く繧ｹ繝・

- [SHP-11] 繝・く繧ｹ繝医↑縺怜峙蠖｢縺ｯ text="" 縺ｫ縺ｪ繧・
- [SHP-12] 隍・焚谿ｵ關ｽ縺ｮ繝・く繧ｹ繝医ｒ謚ｽ蜃ｺ蜿ｯ閭ｽ

---

## **2.3 Arrow & Direction Deduction Requirements**

遏｢蜊ｰ蝗ｳ蠖｢縺ｮ譁ｹ蜷第耳螳壹・邊ｾ蠎ｦ隕∽ｻｶ縲・

- [DIR-01] 0ﾂｰ ﾂｱ22.5ﾂｰ 竊・"E"
- [DIR-02] 45ﾂｰ ﾂｱ22.5ﾂｰ 竊・"NE"
- [DIR-03] 90ﾂｰ ﾂｱ22.5ﾂｰ 竊・"N"
- [DIR-04] 135ﾂｰ ﾂｱ22.5ﾂｰ 竊・"NW"
- [DIR-05] 180ﾂｰ ﾂｱ22.5ﾂｰ 竊・"W"
- [DIR-06] 225ﾂｰ ﾂｱ22.5ﾂｰ 竊・"SW"
- [DIR-07] 270ﾂｰ ﾂｱ22.5ﾂｰ 竊・"S"
- [DIR-08] 315ﾂｰ ﾂｱ22.5ﾂｰ 竊・"SE"
- [DIR-09] 蠅・阜隗貞ｺｦ縺ｮ蝣ｴ蜷医∫援蛛ｴ縺ｫ荳ｸ繧√ｋ・井ｻ墓ｧ倥←縺翫ｊ・・

---

## **2.4 Chart Extraction Requirements**

### Chart meta

- [CH-01] ChartType 縺・XL_CHART_TYPE_MAP 縺ｫ蝓ｺ縺･縺肴枚蟄怜・蛹悶＆繧後ｋ
- [CH-02] Chart Title 縺悟叙蠕励＆繧後ｋ・医↑縺・ｴ蜷医・ null・・
- [CH-03] y_axis_title 縺梧ｭ｣縺励￥蜿門ｾ励＆繧後ｋ・医↑縺・ｴ蜷医・遨ｺ譁・ｭ暦ｼ・

### Axis range

- [CH-04] 譛蟆・譛螟ｧ蛟､縺・float 縺ｧ蜿門ｾ励＆繧後ｋ
- [CH-05] 譛ｪ險ｭ螳壽凾縺ｯ遨ｺ list 繧定ｿ斐☆

### Series meta

- [CH-06] name_range 縺・Excel 蜿ら・蠑上〒蜃ｺ蜉帙＆繧後ｋ・井ｾ・ =Sheet1!$B$1・・
- [CH-07] x_range 縺悟盾辣ｧ蠑上〒蜃ｺ蜉帙＆繧後ｋ
- [CH-08] y_range 縺悟盾辣ｧ蠑上〒蜃ｺ蜉帙＆繧後ｋ
- [CH-09] 謨｣蟶・峙, 蜀・げ繝ｩ繝・ 譽偵げ繝ｩ繝輔↑縺ｩ蜈ｨ繧ｿ繧､繝励′隗｣譫先・蜉溘☆繧・

### 繧ｨ繝ｩ繝ｼ蜃ｦ逅・

- [CH-10] 隗｣譫仙､ｱ謨玲凾 error 縺ｫ繝｡繝・そ繝ｼ繧ｸ縺悟・繧翫け繝ｩ繝・す繝･縺励↑縺・

---

## **2.5 Layout Integration Requirements**

蝗ｳ蠖｢縺ｨ繧ｻ繝ｫ縺ｮ諢丞袖逧・ｴ舌▼縺代↓髢｢縺吶ｋ隕∽ｻｶ縲・

- [LAY-01] Shape 縺ｮ荳ｭ蠢・せ縺悟ｱ槭☆繧玖｡・r 繧呈ｭ｣縺励￥謗ｨ螳壹〒縺阪ｋ
- [LAY-02] 蛻玲婿蜷代・邏舌▼縺代・莉墓ｧ倥↓蠕薙＞邁｡譏薙↓陦後≧・域悴螳溯｣・↑繧・test skip・・
- [LAY-03] 1 陦後↓隍・焚縺ｮ shapes 縺御ｻ倥￥蝣ｴ蜷・shape 鬆・ｺ上ｒ菫晄戟縺吶ｋ
- [LAY-04] 繧ｷ繝ｼ繝医↓ shapes 縺後↑縺・ｴ蜷医・遨ｺ list

---

# 3. Model Validation Requirements

pydantic 讒矩縺悟ｿ・★莉墓ｧ倥←縺翫ｊ縺ｧ縺ゅｋ縺薙→繧呈､懆ｨｼ縺吶ｋ縲・

- [MOD-01] 縺吶∋縺ｦ縺ｮ繝｢繝・Ν縺・`BaseModel` 繧堤ｶ呎価縺励※縺・ｋ
- [MOD-02] 蝙九′ DATA_MODEL.md 縺ｫ螳悟・荳閾ｴ縺吶ｋ
- [MOD-03] Optional 縺ｮ鬆・岼縺ｯ譛ｪ謖・ｮ壹〒 None 縺ｫ縺ｪ繧・
- [MOD-04] 謨ｰ蛟､鬆・岼縺ｯ int/float 縺ｨ縺励※豁｣隕丞喧縺輔ｌ繧・
- [MOD-05] direction 縺ｮ Literal 縺御ｻ墓ｧ伜､悶・蝣ｴ蜷・ValidationError 繧呈兜縺偵ｋ
- [MOD-06] rows/shapes/charts/tables 縺後ョ繝輔か繝ｫ繝医〒遨ｺ list 縺ｫ縺ｪ繧・
- [MOD-07] WorkbookData 縺ｯ `__getitem__` 縺ｧ繧ｷ繝ｼ繝亥錐謖・ｮ壹・蜿門ｾ励′縺ｧ縺阪～__iter__` 縺ｧ (sheet_name, SheetData) 繧帝・ｺ冗ｶｭ謖√〒襍ｰ譟ｻ縺ｧ縺阪ｋ

---

# 4. Export Requirements (JSON/YAML/TOON)

- [EXP-01] 空値（None, "", [], {}）は dict_without_empty_values により除外される
- [EXP-02] JSON 出力が UTF-8 で行われる
- [EXP-03] YAML 出力が sort_keys=False で行われる
- [EXP-04] TOON 出力が正しく生成される
- [EXP-05] WorkbookData → JSON → WorkbookData の round-trip が破壊的変更にならない
- [EXP-06] export_sheets でシートごとにファイルが出力される
- [EXP-07] WorkbookData/SheetData の `to_json` は pretty オプションでインデントされる
- [EXP-08] WorkbookData/SheetData の `save(path)` が拡張子でフォーマットを自動判別し、未対応拡張子は ValueError となる
- [EXP-09] WorkbookData/SheetData の `to_yaml` / `to_toon` は依存未導入時に MissingDependencyError を返し、導入済みなら正常に返す
- [EXP-10] ExStructEngine の OutputOptions で include_shapes/charts/tables/rows を False にすると対応フィールドが出力から除外される（空リストも消える）
- [EXP-11] print_areas_dir / save_print_area_views で PrintArea ごとのファイルを出力でき、印刷範囲が無い場合は何も書き出さない
- [EXP-12] PrintAreaView は area に完全に含まれる行のみを保持し、範囲外のセル・リンクを落とす
- [EXP-13] PrintAreaView の table_candidates は印刷範囲に完全に収まる候補のみを保持する
- [EXP-14] normalize オプション指定時、PrintAreaView の行・列インデックスは印刷範囲起点に再基準化される
- [EXP-15] OutputOptions.include_print_areas=False のとき、print_areas_dir が指定されても印刷範囲ファイルを出力しない
- [EXP-16] PrintAreaView に shapes/charts を含め、印刷範囲と交差するもののみ出力する（サイズ不明の図形は点として判定）
- [EXP-17] Chart.w/h は verbose では出力され、standard ではデフォルト出力しない（include_chart_size フラグで制御）
- [EXP-18] Shape.w/h の出力は include_shape_size フラグで制御され、デフォルトは verbose のみ True
- [EXP-19] auto_page_breaks_dir を指定し、ExStructEngine では extract_workbook に include_auto_page_breaks=True が渡り、auto_print_areas が取得される（COM 環境前提）
- [EXP-20] export_auto_page_breaks は auto_print_areas が空の場合に明示的な例外を返し、存在する場合のみ書き出す
- [EXP-21] save_auto_page_break_views で auto_print_areas が指定パスに書き出され、キーに Sheet1#auto#1 などが一意に付与される
- [EXP-22] serialize_workbook に未対応フォーマットを渡すと SerializationError が発生する
# 5. CLI Requirements

- [CLI-01] `exstruct extract file.xlsx` 縺梧・蜉溘☆繧・
- [CLI-02] `--format json/yaml/toml` 縺梧ｩ溯・縺吶ｋ
- [CLI-03] `--image` 縺ｧ PNG 縺悟・蜉帙＆繧後ｋ
- [CLI-04] `--pdf` 縺ｧ PDF 縺悟・蜉帙＆繧後ｋ
- [CLI-05] 辟｡蜉ｹ繝輔ぃ繧､繝ｫ驕ｸ謚樊凾縺ｯ螳牙・縺ｫ邨ゆｺ・☆繧・
- [CLI-06] 繧ｨ繝ｩ繝ｼ繝｡繝・そ繝ｼ繧ｸ縺・stdout 縺ｫ蜃ｺ蜉帙＆繧後ｋ
- [CLI-07] `--print-areas-dir` 謖・ｮ壽凾縺ｫ蜊ｰ蛻ｷ遽・峇縺斐→縺ｮ繝輔ぃ繧､繝ｫ縺悟・蜉帙＆繧後ｋ・・nclude_print_areas=False 縺ｮ蝣ｴ蜷医・繧ｹ繧ｭ繝・・・・

---

# 6. Error Handling Requirements

- [ERR-01] xlwings COM エラーでもプロセスが落ちない
- [ERR-02] 図形抽出失敗時でも他要素が取得される
- [ERR-03] Chart extraction failure は Chart.error に明示される
- [ERR-04] 異常な参照範囲（broken range）は例外化せず null を error に記録
- [ERR-05] Excel ファイルが開けない場合にメッセージを出して終了する
- [ERR-06] openpyxl の `_print_area` に設定された印刷範囲が存在する場合でも抽出漏れしない
- [ERR-07] export_auto_page_breaks は auto_print_areas が空の場合に PrintAreaError（ValueError 互換）を送出する
- [ERR-08] YAML/TOON 依存が未導入の場合、MissingDependencyError が発生しインストール手順を案内する
- [ERR-09] ファイル書き込みに失敗した場合、OutputError が送出される（元例外は __cause__ に保持）
# 7. Regression Requirements

- [REG-01] 驕主悉繝舌・繧ｸ繝ｧ繝ｳ縺ｨ蜷後§ Excel 繧貞・蜉帙＠縺溘→縺阪∝・蜉・JSON 縺ｮ讒矩縺悟､峨ｏ繧峨↑縺・
- [REG-02] Models 縺ｮ繧ｭ繝ｼ蜑企勁 or 蜷榊燕螟画峩縺ｯ縺吶∋縺ｦ遐ｴ螢顔噪螟画峩縺ｨ縺励※讀懃衍縺吶ｋ
- [REG-03] 譁ｹ蜷第耳螳壹い繝ｫ繧ｴ繝ｪ繧ｺ繝縺ｮ螟画峩讀懃衍
- [REG-04] ChartSeries 縺ｮ蜿ら・遽・峇隗｣譫舌′驕主悉邨先棡縺ｨ荳閾ｴ縺吶ｋ

---

# 8. Non-Functional Requirements

### Performance

<!-- - [PERF-01] 譛ｪ螳・-->

### Memory

<!-- - [MEM-01] 100MB 縺ｮ Excel 繧呈桶縺・圀縺ｫ Python 繝励Ο繧ｻ繧ｹ縺・1GB 繧定ｶ・∴縺ｪ縺・
- [MEM-02] 繝ｬ繝ｳ繝繝ｪ繝ｳ繧ｰ・・NG・画凾縺ｫ繝ｪ繝ｼ繧ｯ縺後↑縺・-->

---

# 9. Mode Output Requirements

- [MODE-01] CLI `--mode` 縺ｨ API `extract(..., mode=)` 縺・`light`/`standard`/`verbose` 縺ｮ縺ｿ蜿励￠莉倥￠縲√ョ繝輔か繝ｫ繝医・ `standard`
- [MODE-02] `light` 繝｢繝ｼ繝峨・繧ｻ繝ｫ縺ｨ繝・・繝悶Ν縺ｮ縺ｿ霑斐＠縲《hapes/charts 縺ｯ遨ｺ縺ｧ COM 繧｢繧ｯ繧ｻ繧ｹ繧ゅ＠縺ｪ縺・
- [MODE-03] `standard` 繝｢繝ｼ繝峨・譌｢蟄俶嫌蜍輔ｒ邯ｭ謖√＠縲√ユ繧ｭ繧ｹ繝井ｻ倥″蝗ｳ蠖｢縺ｾ縺溘・遏｢蜊ｰ邉ｻ縺ｮ縺ｿ蜃ｺ蜉帙＠縲，OM 譛牙柑譎ゅ・繝√Ε繝ｼ繝亥叙蠕・
- [MODE-04] `verbose` 繝｢繝ｼ繝峨・ chart/comment/picture/form control 莉･螟悶・蜈ｨ蝗ｳ蠖｢繧貞・蜉帙＠縲√ユ繧ｭ繧ｹ繝医・譛臥┌縺ｫ縺九°繧上ｉ縺・`w`/`h` 繧貞ｿ・★蜷ｫ繧√ｋ
- [MODE-05] `process_excel` 縺ｧ繝｢繝ｼ繝画欠螳壹′莨晄成縺励￣DF/逕ｻ蜒上が繝励す繝ｧ繝ｳ菴ｵ逕ｨ縺ｧ繧よｭ｣蟶ｸ邨ゆｺ・☆繧・
- [MODE-06] `standard` 繝｢繝ｼ繝峨〒譌｢蟄倥ヵ繧｣繧ｯ繧ｹ繝√Ε縺ｮ蜃ｺ蜉帙↓蝗槫ｸｰ縺後↑縺・ｼ井ｸ崎ｦ√↑蝗ｳ蠖｢縺悟｢励∴縺ｪ縺・ｼ・
- [MODE-07] 辟｡蜉ｹ縺ｪ繝｢繝ｼ繝牙､縺ｯ蜃ｦ逅・幕蟋句燕縺ｫ繧ｨ繝ｩ繝ｼ縺ｨ縺ｪ繧・
- [INT-01] COM 繧ｪ繝ｼ繝励Φ螟ｱ謨玲凾縺ｫ `extract_workbook` 縺後そ繝ｫ・九ユ繝ｼ繝悶Ν蛟呵｣懊・縺ｿ繧定ｿ斐☆繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ繧定｡後≧
- [IO-05] `dict_without_empty_values` 縺・None/遨ｺ繝ｪ繧ｹ繝・遨ｺ霎樊嶌/遨ｺ譁・ｭ怜・繧帝勁蜴ｻ縺励ロ繧ｹ繝域ｧ矩繧剃ｿ晄戟縺吶ｋ
- [RENDER-01] Excel+COM+pypdfium2 迺ｰ蠅・〒 PDF/PNG 繧貞・蜉帙〒縺阪ｋ繧ｹ繝｢繝ｼ繧ｯ繝・せ繝茨ｼ育腸蠅・､画焚縺ｧ繧ｪ繝ｳ繧ｪ繝募庄閭ｽ・・
- [MODE-08] `light` 繝｢繝ｼ繝峨〒繧ょ魂蛻ｷ遽・峇繧・openpyxl 縺ｧ謚ｽ蜃ｺ縺吶ｋ縺後√ョ繝輔か繝ｫ繝亥・蜉帙〒縺ｯ print_areas 繧貞性繧√↑縺・ｼ・uto 蛻､螳夲ｼ・


