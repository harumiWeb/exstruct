# ExStruct Coding Guidelines

**AI（Codex）と人間が協調して長期的に保守できるコードベース** を実現するための、  
ExStruct 独自のコーディング規約です。

以下の規約は、Codex にコードを生成させる場合にも適用します。  
AI 生成コードの品質安定化・保守性向上・Ruff / mypy / Pydantic との整合性を目的とします。

---

# 1. 基本方針（最重要）

## 1.1 型ヒントは必須

すべての関数・メソッド・クラスに完全な型ヒントを付けること。

**良い例**

```python
def extract_shapes(sheet: xw.Sheet) -> list[Shape]:
    ...
```

**悪い例**

```python
def extract_shapes(sheet):
    ...
```

---

## 1.2 関数は「1 責務」だけを持つ

複雑な Excel 処理は関数が肥大化しやすいため、
**1 関数 = 1 ロジック** を徹底する。

**良い例**

- `extract_raw_shapes`
- `normalize_shape`
- `detect_shape_direction`

---

## 1.3 境界は BaseModel、内部は dataclass を主なデータ構造にする

Raw*Data (e.g., SheetRawData/WorkbookRawData) は内部専用で公開APIに出さない。
Do not expose them via public APIs or re-export from package __init__.

辞書やタプルではなく、**必ず Pydantic BaseModel か dataclass を返す**。

メリット：

- AI が誤解しにくい
- IDE での開発が容易
- JSON 化との相性がよい

---

# 2. 命名規則

## 2.1 snake_case（関数・変数）

例：`parse_chart_labels`, `shape_items`

## 2.2 PascalCase（クラス）

例：`WorkbookParser`, `ChartSeries`

## 2.3 モジュール名は短く責務を示す

例：`shape_parser.py`, `chart_reader.py`

---

# 3. インポート規則

Ruff の isort ルールに従う。

順序：

1. 標準ライブラリ
2. サードパーティ
3. 自前ライブラリ（exstruct）

**例**

```python
import json
from typing import Any

import xlwings as xw
from pydantic import BaseModel

from exstruct.models import Shape
```

---

# 4. Docstring 規約

Google スタイルを採用し、下記を必ず明記：

- Args
- Returns
- Raises（例外を投げる場合）

**例**

```python
def detect_shape_direction(shape: Shape) -> str | None:
    """Detect arrow direction from shape coordinates and rotation.

    Args:
        shape: Parsed shape model.

    Returns:
        Direction code ("E", "NE", etc.) or None if no arrow is detected.
    """
```

---

# 5. 例外処理のルール

## 5.1 ValueError / RuntimeError を中心に使用

```python
if not cell:
    raise ValueError("cell must not be empty.")
```

## 5.2 COM 例外はラップする

```python
try:
    text = shape.text_frame.characters.text
except Exception as e:
    raise RuntimeError(f"Failed to read shape text: {e}") from e
```

---

# 6. 複雑度のコントロール

Ruff の `C90`（mccabe）に従い、
**max-complexity = 12** を超えないように関数を分割する。

---

# 7. Codex（AI）を使うときのガイドライン

## 7.1 Codex に守らせるルール

Codex を利用する場合、以下を事前にプロンプトに記載する：

- 型ヒントは必須
- 1 関数 = 1 責務
- 境界は BaseModel、内部は dataclass モデルを返す
- Docstring を Google スタイルで書く
- import は正しい順序で書く
- エラーハンドリングは簡潔に
- 複雑度が高くなる場合は関数を分割する

---

## 7.2 Codex にアウトプットさせるための推奨プロンプト

以下を貼るだけで安定したコードを生成できる：

```
あなたは熟練 Python エンジニアであり、ExStruct ライブラリの標準に従ってコードを書きます。

必ず以下を守ってください：
- 型ヒントは全ての引数と戻り値に付ける
- １関数 = １責務
- 境界は BaseModel、内部は dataclass を返す
- import を正しい順序で並べる
- docstring（Google スタイル）を書く
- 複雑になりすぎないよう関数を分割する
- JSON や辞書ではなく Pydantic モデルを返す

出力は Python のコードのみ。
```

---

# 8. AI 生成コードのレビュー項目（チェックリスト）

AI が生成したコードをレビューする際は、以下の順に確認する：

1. 型ヒントが完全か
2. docstring があるか
3. import の順序が正しいか
4. 境界は BaseModel、内部は dataclass モデルを返しているか
5. 関数が単一責務か
6. 例外処理が適切か
7. 複雑度が max 12 を超えていないか
8. Ruff でエラーが出ないか

---

# 9. 禁止事項（AI 生成コードで特に注意）

以下は Codex に絶対させない：

- 複雑すぎる if/else ネスト
- 多数の責務を持つ「神クラス」
- 大きな辞書やタプルを返す
- コメントが一切ないコード
- 無名のマジックナンバー

---

# 10. ExStruct の設計原則まとめ

- 型安全性（type safety）
- データ構造の明確化（境界は BaseModel、内部は dataclass）
- モジュールの疎結合化
- 関数責務の最小化
- AI と人間の共存を前提にしたコード規律

これらは **Ruff / mypy / CI とも高い互換性**を持つよう設計されている。

---

# Appendix: 今後追加が推奨されるルール

- mypy の厳格化
- pydantic のフィールド制約統一
- public API の docstring 強制
- internal API の \_prefix 命名
