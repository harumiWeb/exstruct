# Task List

未完了 [ ], 完了 [x]

- [x] 仕様: `merged_cells` の新フォーマット（schema + items）をモデルと出力仕様に反映
- [x] 仕様: `include_merged_values_in_rows` フラグ追加（デフォルト True）
- [x] 実装: 既存の `merged_cells` 生成ロジックを新構造へ置換
- [x] 実装: `rows` から結合セル値を排除する分岐を追加（フラグ制御）
- [x] 実装: 結合セルの値がない場合は `" "` を出力
- [ ] 更新: 既存の JSON 出力例・ドキュメントの整合性確認
- [x] テスト: 結合セルが多いケースの JSON 量削減を確認
