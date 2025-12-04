# Feature Spec (今後実装予定の仕様)

- 出力情報量を("light", "standard", "verbose")の 3 つから選択できるように
  - デフォルトは standard (セル、グラフ、テキストのある図形と矢印系の図形、表構造検出)
  - COM 無し環境や軽量データなら light（セル、表構造検出のみ）
  - verbose は取得できるデータ全てを出力
