# CONTRIBUTING FOR AI AGENTS

このファイルは AI コーディングエージェント（ChatGPT, Cursor, Copilot 等）向けの  
**特別なガイドライン**です。

## 原則

1. **docs/\*.md を仕様の唯一のソースとする**
2. モデル定義（DATA_MODEL.md）と矛盾するコードは禁止
3. core 層は「抽出のみ」。推論ロジックは integrate に集約
4. models 層は絶対に「副作用なし」
5. I/O 処理とコアロジックを混在させない
6. 例外処理は fail-safe を徹底する
7. 新機能を追加するときはロードマップを更新すること

## AI 向けタスクセパレーション

- 新しい抽出機能, 意味解析アルゴリズム → core/
- 新しいデータ構造 → models/
- 出力形式追加 → io/
- CLI 機能 → cli/

## テスト方針

- テスティングフレームワークは`pytest`, `pytest-mock`を使用
- Excel ファイルのサンプルは `/tests/data/*.xlsx` に置く
- 回帰テストとして Pydantic モデル一致を優先する
