# Task List

## 結合セル（MergedCell）テスト

- [ ] フィクスチャ作成：結合セルあり/なし、複数範囲、値あり/空文字の Excel を用意
- [ ] OpenpyxlBackend.extract_merged_cells の正常系（座標・代表値）をユニットテスト
- [ ] OpenpyxlBackend.extract_merged_cells の例外時フォールバック（空マップ）をテスト
- [ ] ComBackend.extract_merged_cells が NotImplementedError を送出することをテスト
- [ ] Pipeline: standard/verbose で merged_cells を含み、light では空になることをテスト
- [ ] Pipeline: include_merged_cells=False で抽出ステップが無効化されることをテスト
- [ ] Modeling: SheetRawData→SheetData で merged_cells が保持されることをテスト
- [ ] Engine: OutputOptions.filters.include_merged_cells=False で出力から除外されることをテスト
- [ ] Export: dict_without_empty_values により merged_cells 空リストが出力されないことをテスト

## カバレッジ対応

- [ ] 追加テストで 78% 以上の全体カバレッジを満たすことを確認