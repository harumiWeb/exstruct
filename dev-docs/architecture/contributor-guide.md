# コントリビューターガイド - 内部アーキテクチャ

## 対象読者

このページは、次の人向けである。

- ExStruct の内部実装を拡張したい人
- 新しい抽出対象（shapes, SmartArt, comments など）を追加したい人
- backend（Openpyxl / COM / LibreOffice / 将来の XML）を拡張したい人
- どこを触るべきか迷いながら PR を出そうとしている人

---

## ディレクトリ構成（core）

```text
src/exstruct/core/
├── pipeline.py        # Orchestrates the overall flow
├── backends/          # Backend abstractions and runtime-specific adapters
│   ├── openpyxl_backend.py
│   ├── com_backend.py
│   └── libreoffice_backend.py
├── libreoffice.py     # LibreOffice runtime/session helper
├── ooxml_drawing.py   # OOXML drawing/chart parser for best-effort rich extraction
├── modeling.py        # Final data integration
├── workbook.py        # Workbook lifecycle management
├── cells.py           # Cell/table analysis (mainly openpyxl)
└── utils.py           # Shared utilities
```

---

## 重要な設計ルール

### 1. Pipeline は順序だけを知る

- Excel parsing logic を Pipeline に置かない
- Pipeline の責務は次だけに限定する
  - Backend の呼び出し順序
  - fallback の判断
  - artifact の管理
  - Modeling への受け渡し

**判断基準**

> このコードは Excel の中身を直接読んでいるか？
> 読んでいるなら Pipeline に置くべきではない。

---

### 2. Backend は抽出専用

Backend は **純粋な抽出** のためにある。

- Excel -> raw data
- 解釈しない
- 統合しない
- 可能な限り副作用を避ける

#### Backend で許可すること

- cell value の読み取り
- shape position の読み取り
- COM API の呼び出し
- 例外の送出

#### Backend で許可しないこと

- WorkbookData / SheetData の構築
- 出力形式の都合を持ち込むこと
- fallback logging（これは Pipeline が担当する）

---

### 3. Modeling を唯一の統合点にする

複数 backend の結果を 1 つの **意味構造** に統合するのは Modeling だけにする。

- Openpyxl + COM / LibreOffice の結果を結合する
- 座標、方向、type を正規化する
- 欠損データを補う

> 最終的な JSON/YAML/TOON の形を知ってよい層は
> **Modeling** だけである。

---

## よくある拡張パターン

---

## ケース 1: 新しい抽出対象を追加する（例: comments）

### 手順

1. **Backend に抽出メソッドを追加する**

   ```python
   class Backend(Protocol):
       def extract_comments(self, ...): ...
   ```

2. `OpenpyxlBackend` / `ComBackend` に実装する
   - 片側だけでもよい。未実装なら `NotImplementedError` を使う

3. `pipeline.py` に呼び出しを追加する
   - fallback 対象に含めるかどうかを明示する

4. `modeling.py` で WorkbookData に統合する

5. test を追加する

---

## ケース 2: 新しい Backend を追加する（例: XML または LibreOffice backend）

### 手順

1. `backend.py` の Protocol を実装する

   ```python
   class XmlBackend:
       def extract_cells(...)
       def extract_shapes(...)
   ```

2. Pipeline に backend 選択を追加する
   - 既存 backend への変更は最小にとどめる

3. 可能な限り Modeling は変えない

---

## ケース 3: 出力構造を変更する

- **これは最も壊れやすい変更種別である**

### 原則

- 変更を `modeling.py` と Pydantic model に限定する
- backend は変えない
- Pipeline は変更しない

---

## フォールバックルール

- COM または LibreOffice runtime の unavailable は **通常系** である
- fallback は例外扱いしない
- 常に `FallbackReason` を与える

```python
log_fallback(
    reason=FallbackReason.COM_UNAVAILABLE,
    message="COM backend not available"
)

log_fallback(
    reason=FallbackReason.LIBREOFFICE_UNAVAILABLE,
    message="LibreOffice backend not available"
)
```

---

## テストガイドライン

### 期待する test 粒度

| Layer    | Test focus           |
| -------- | -------------------- |
| Backend  | extraction correctness |
| Pipeline | fallback / branching |
| Modeling | integration logic    |

### アンチパターン

- 実 Excel instance に強く依存する fragile test
- Backend と Modeling を一気に結合する巨大 test

---

## PR 前チェックリスト

- [ ] Pipeline に Excel parsing logic が入っていない
- [ ] Backend に解釈 logic が入っていない
- [ ] 最終構造の single source が Modeling になっている
- [ ] Fallback reason が明示されている
- [ ] Tests が追加されている
- [ ] Public API を変えたなら docs が更新されている

---

## よくあるアンチパターン

- Backend の中で WorkbookData を作る
- Pipeline から openpyxl / xlwings を直接呼ぶ
- "とりあえずここで処理する" 型の ad-hoc logic
- fallback reason のない catch-all exception

---

## 設計思想の要約

- Excel は **fragile**
- COM は **powerful but unstable**
- LLM/RAG は **stable structure first** を要求する

そのため、

> 責務を分離し、失敗点を局所化する。
