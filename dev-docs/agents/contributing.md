# AI エージェント向けコントリビュートガイド

このファイルは AI コーディングエージェント（ChatGPT, Cursor, Copilot 等）向けの  
**特別なガイドライン**です。

## 原則

1. `docs/` は公開契約、`dev-docs/specs/` は内部仕様、`dev-docs/adr/` は判断理由として読む
2. モデル定義（`dev-docs/specs/data-model.md`）と矛盾するコードは禁止
3. core 層は「抽出のみ」。統合ロジックは `modeling.py` に集約し、`integrate.py` は pipeline 呼び出しの入口にとどめる
4. models 層は絶対に「副作用なし」
5. I/O 処理とコアロジックを混在させない
6. 例外処理は fail-safe を徹底する
7. 新機能を追加するときはロードマップを更新すること

## 参照優先順位

1. `docs/`
2. `dev-docs/specs/`
3. `dev-docs/adr/`
4. `tests/`
5. `src/`

役割分担:

- ADR = なぜそうしたか
- specs = 何を保証するか
- tests = その振る舞いの証拠
- src = どう実装しているか

## AI 向けタスクセパレーション

- 新しい抽出機能, 意味解析アルゴリズム → core/
- 新しいデータ構造 → models/
- 出力形式追加 → io/
- CLI 機能 → cli/

## コーディングガイドライン

必ず以下を守ってください：

- 型ヒントは全ての引数と戻り値に付ける
- １関数 = １責務
- 境界は BaseModel、内部は dataclass を返す
- import を正しい順序で並べる
- docstring（Google スタイル）を書く
- 複雑になりすぎないよう関数を分割する
- JSON や辞書ではなく Pydantic モデルを返す

## テスト方針

- テスティングフレームワークは`pytest`, `pytest-mock`を使用
- Excel ファイルのサンプルは `/tests/data/*.xlsx` に置く
- 回帰テストとして Pydantic/dataclass モデル一致を優先する
- 静的解析にはruff, mypyを使用する。したがって、これらのlintに通る実装をする。
