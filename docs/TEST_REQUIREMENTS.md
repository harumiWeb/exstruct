# ExStruct Test Requirements Specification

Version: 0.2  
Status: Required for Release

この文書は、ExStruct のすべての機能に対する **正式なテスト要件一覧** であり、  
AI エージェント・人間開発者双方が参照して自動/手動テストを生成できるように設計されています。

---

# 1. Test Coverage Categories

ExStruct のテストは以下のカテゴリに分類される：

1. **セル抽出（Cells Extraction）**
2. **図形抽出（Shapes Extraction）**
3. **矢印・方向推定（Arrow + Direction Detection）**
4. **チャート抽出（Chart Extraction）**
5. **意味付与（Layout Integration）**
6. **データモデル準拠テスト（Pydantic Validation）**
7. **出力フォーマット（JSON/YAML/TOML Writer）**
8. **CLI テスト**
9. **エラー処理・フェイルセーフ**
10. **回帰テスト（Regression）**
11. **パフォーマンス/メモリ要件**

---

# 2. Functional Test Requirements (詳細要件)

---

## **2.1 Cells Extraction Requirements**

### 必須テスト

- [CEL-01] 空セルを除外し、非空セルのみ `c` に出力される
- [CEL-02] 行番号 `r` が 0-based index で正しく出力される
- [CEL-03] 列番号が `"0"`, `"1"` の **文字列キー** で出力される
- [CEL-04] セルに改行・タブが含まれても正しく読み込める
- [CEL-05] Unicode（絵文字、日本語、異体字）セルの読み取り
- [CEL-06] Pandas 読み込みによる dtype=string 強制が守られている
- [CEL-07] セル範囲が大きいファイルでも 1 万セル程度で性能問題がない

---

## **2.2 Shapes Extraction Requirements**

### 基本形状

- [SHP-01] AutoShape の type が正しく文字列化される
- [SHP-02] TextFrame の文字列が正しく読み取れる
- [SHP-03] サイズ（w, h）が null にならない（取得不可時は null）
- [SHP-04] Group（グループ図形）内の子図形がすべて展開される or 無視方針が維持される

### 座標

- [SHP-05] l, t（left, top）が整数で取得される
- [SHP-06] 表示倍率やウィンドウズームが変わっても座標が変動しない

### 回転 / 矢印

- [SHP-07] rotation が Excel の回転角度値に一致する
- [SHP-08] angle_deg が線図形の向きとして正しく算出される
- [SHP-09] begin_arrow_style / end_arrow_style が Excel の ENUM と一致する
- [SHP-10] direction が 8 方位分類に従い正しく算出される

### テキスト

- [SHP-11] テキストなし図形は text="" になる
- [SHP-12] 複数段落のテキストを抽出可能

---

## **2.3 Arrow & Direction Deduction Requirements**

矢印図形の方向推定の精度要件。

- [DIR-01] 0° ±22.5° → "E"
- [DIR-02] 45° ±22.5° → "NE"
- [DIR-03] 90° ±22.5° → "N"
- [DIR-04] 135° ±22.5° → "NW"
- [DIR-05] 180° ±22.5° → "W"
- [DIR-06] 225° ±22.5° → "SW"
- [DIR-07] 270° ±22.5° → "S"
- [DIR-08] 315° ±22.5° → "SE"
- [DIR-09] 境界角度の場合、片側に丸める（仕様どおり）

---

## **2.4 Chart Extraction Requirements**

### Chart meta

- [CH-01] ChartType が XL_CHART_TYPE_MAP に基づき文字列化される
- [CH-02] Chart Title が取得される（ない場合は null）
- [CH-03] y_axis_title が正しく取得される（ない場合は空文字）

### Axis range

- [CH-04] 最小/最大値が float で取得される
- [CH-05] 未設定時は空 list を返す

### Series meta

- [CH-06] name_range が Excel 参照式で出力される（例: =Sheet1!$B$1）
- [CH-07] x_range が参照式で出力される
- [CH-08] y_range が参照式で出力される
- [CH-09] 散布図, 円グラフ, 棒グラフなど全タイプが解析成功する

### エラー処理

- [CH-10] 解析失敗時 error にメッセージが入りクラッシュしない

---

## **2.5 Layout Integration Requirements**

図形とセルの意味的紐づけに関する要件。

- [LAY-01] Shape の中心点が属する行 r を正しく推定できる
- [LAY-02] 列方向の紐づけは仕様に従い簡易に行う（未実装なら test skip）
- [LAY-03] 1 行に複数の shapes が付く場合 shape 順序を保持する
- [LAY-04] シートに shapes がない場合は空 list

---

# 3. Model Validation Requirements

pydantic 構造が必ず仕様どおりであることを検証する。

- [MOD-01] すべてのモデルが `BaseModel` を継承している
- [MOD-02] 型が DATA_MODEL.md に完全一致する
- [MOD-03] Optional の項目は未指定で None になる
- [MOD-04] 数値項目は int/float として正規化される
- [MOD-05] direction の Literal が仕様外の場合 ValidationError を投げる
- [MOD-06] rows/shapes/charts/tables がデフォルトで空 list になる

---

# 4. Export Requirements（JSON/YAML/TOML）

- [EXP-01] 空値（None, "", [], {}）は dict_without_empty_values により除外される
- [EXP-02] JSON 出力が UTF-8 で行われる
- [EXP-03] YAML 出力が sort_keys=False で行われる
- [EXP-04] TOON 出力がバイナリ書き込みで正しく生成される
- [EXP-05] WorkbookData → JSON → WorkbookData の round-trip が破壊的変更なし
- [EXP-06] export_sheets でシートごとにファイルが出力される

---

# 5. CLI Requirements

- [CLI-01] `exstruct extract file.xlsx` が成功する
- [CLI-02] `--format json/yaml/toml` が機能する
- [CLI-03] `--image` で PNG が出力される
- [CLI-04] `--pdf` で PDF が出力される
- [CLI-05] 無効ファイル選択時は安全に終了する
- [CLI-06] エラーメッセージが stdout に出力される

---

# 6. Error Handling Requirements

- [ERR-01] xlwings COM エラーでもプロセスが落ちない
- [ERR-02] 図形抽出失敗時でも他要素が取得される
- [ERR-03] Chart extraction failure → Chart.error に必ず文字列
- [ERR-04] 異常な参照（broken range）は例外化せず null か error に記録
- [ERR-05] Excel ファイルが開けない場合はメッセージを出して終了

---

# 7. Regression Requirements

- [REG-01] 過去バージョンと同じ Excel を入力したとき、出力 JSON の構造が変わらない
- [REG-02] Models のキー削除 or 名前変更はすべて破壊的変更として検知する
- [REG-03] 方向推定アルゴリズムの変更検知
- [REG-04] ChartSeries の参照範囲解析が過去結果と一致する

---

# 8. Non-Functional Requirements

### Performance

<!-- - [PERF-01] 未定 -->

### Memory

- [MEM-01] 100MB の Excel を扱う際に Python プロセスが 1GB を超えない
- [MEM-02] レンダリング（PNG）時にリークがない

---

# 9. Mode Output Requirements

- [MODE-01] CLI `--mode` と API `extract(..., mode=)` が `light`/`standard`/`verbose` のみ受け付け、デフォルトは `standard`
- [MODE-02] `light` モードはセルとテーブルのみ返し、shapes/charts は空で COM アクセスもしない
- [MODE-03] `standard` モードは既存挙動を維持し、テキスト付き図形または矢印系のみ出力し、COM 有効時はチャート取得
- [MODE-04] `verbose` モードは chart/comment/picture/form control 以外の全図形を出力し、テキストの有無にかかわらず `w`/`h` を必ず含める
- [MODE-05] `process_excel` でモード指定が伝搬し、PDF/画像オプション併用でも正常終了する
- [MODE-06] `standard` モードで既存フィクスチャの出力に回帰がない（不要な図形が増えない）
- [MODE-07] 無効なモード値は処理開始前にエラーとなる
