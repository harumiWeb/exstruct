# Task List

## 1. 既存実装の修正（モデル分離の影響対応）
- [x] `src/exstruct/io/__init__.py` の `_filter_shapes_to_area` が `list[Shape | Arrow | SmartArt]` を受け取れるように型と処理を調整する
- [x] `src/exstruct/core/shapes.py` のコネクタ判定を `Arrow` 前提に変更する（`begin_arrow_style` / `end_arrow_style` などは `Arrow` のみ参照）
- [x] `src/exstruct/core/shapes.py` の接続 ID 参照を `Arrow` に限定し、`Shape` からの誤参照を除去する
- [x] `PrintAreaView` 側の `shapes` フィルタで `SmartArt` を落とさないことを確認する

## 2. SmartArt 取得機能の実装方針
- [x] `shape.HasSmartArt` を条件に SmartArt を抽出する
- [x] `SmartArt.Layout.Name` を `SmartArt.layout_name` に格納する
- [x] `SmartArt.AllNodes` を走査し、`level` と `text` を収集する
- [x] ノード配列から `SmartArtNode` のツリー（`roots`）を構築する（`level` を使ったスタック組み立て）
- [x] `SmartArt` は `BaseShape` 相当の位置/サイズ/回転/テキストを併せて格納する

## 3. 実装箇所の整理
- [x] `src/exstruct/core/shapes.py` に SmartArt 抽出用の関数を追加する（1 関数=1 責務を遵守）
- [x] `src/exstruct/core/shapes.py` のメイン抽出処理で `Shape` / `Arrow` / `SmartArt` に振り分ける
- [x] `src/exstruct/io/__init__.py` で `Shape | Arrow | SmartArt` のシリアライズ挙動が崩れないことを確認する

## 4. 動作確認
- [x] 既存の shape / connector 抽出が壊れていないことを確認する
- [ ] SmartArt が含まれるブックで `SmartArt.roots` が期待どおりに出力されることを確認する

## 5. テストケース（カバレッジ維持）
- [x] `SmartArt` の `roots` がネスト構造でシリアライズされることを確認する
- [x] `Arrow` のみが `begin_id` / `end_id` を持ち、`Shape` では参照されないことを確認する
- [x] `_filter_shapes_to_area` が `Shape | Arrow | SmartArt` を受け取り、SmartArt も対象に含めることを確認する
- [x] `kind` による判別が想定どおり動くことを確認する
