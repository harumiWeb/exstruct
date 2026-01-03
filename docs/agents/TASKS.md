# Task List

## セル結合範囲取得（MergedCell）

- [ ] 既存の抽出フローで結合セル情報の入口を確認（Openpyxl/COMの責務境界を整理）
- [ ] Backend Protocol に merged_cells 抽出メソッドを追加（未実装は NotImplementedError）
- [ ] OpenpyxlBackend で結合セル範囲と代表値を抽出（standard/verbose のみ）
- [ ] ComBackend で取得可否を判断し、可能なら実装・難しければ未実装扱い
- [ ] Pipeline に抽出呼び出しを追加（fallback の理由ログを明示）
- [ ] Modeling で raw 結合セル情報を MergedCell に正規化し、SheetData.merged_cells に統合
- [ ] Engine の無効化オプション追加（出力時は OutputOptions で merged_cells を削除）
- [ ] テスト追加（Backend:抽出, Pipeline:fallback, Modeling:統合）
- [ ] ドキュメント更新確認（DATA_MODEL 反映済みか確認）