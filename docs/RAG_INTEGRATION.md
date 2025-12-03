# RAG Integration — How to Use ExStruct in AI Pipelines

ExStruct の想定用途は RAG（Retrieval Augmented Generation）です。

## 推奨フロー

```txt
Excel → ExStruct(JSON) → Chunking → VectorDB → LLM
```

## JSON そのまま格納（構造優先）

- ChartType / AutoShapeType / Coordinates が強い検索キー
- テーブル検索が精確
- Lookup 型 RAG に最適

## Markdown 化して格納（生成品質優先）

- LLM が読みやすい
- 回答が自然になる

## ハイブリッド設計（最も推奨）

```txt
VectorDB #1: Structural JSON
VectorDB #2: Markdown Summary
```

RAG 実行時：

- #1 で構造マッチ
- #2 で自然言語マッチ

両方合わせて LLM に送る。
