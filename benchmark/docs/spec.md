# Reconstruction Utility Benchmark (RUB) Specification

## 0. Scope (v0.1 / lite vs v1)

RUB lite (v0.1) is a small, fast-running subset intended for quick checks.
The full RUB (v1) is the primary benchmark for public reporting.

RUB lite assets:

- benchmark/rub/manifest_lite.json
- benchmark/rub/truth_lite/*.json

Full RUB assets:

- benchmark/rub/manifest.json
- benchmark/rub/truth/*.json

## 1. Goal

RUB measures how useful reconstructed Markdown is for downstream structure-aware
queries. The target is reconstruction utility rather than raw string similarity.

## 2. Inputs and outputs

- Input: Excel workbooks (.xlsx)
- Methods: pdf, image_vlm, exstruct, html, openpyxl
- Stage A output: reconstructed Markdown
- Stage B output: JSON-only answers to structure queries

## 3. Two-stage evaluation

### Stage A: Reconstruction

Each method produces Markdown from the same source workbook.

- pdf: soffice -> pdf -> text extraction -> Markdown
- image_vlm: render -> VLM -> Markdown
- exstruct: exstruct JSON -> LLM -> Markdown
- html / openpyxl: rule-based extraction -> Markdown

### Stage B: Structure queries

Only the Stage A Markdown is used as input to answer queries.

- Output must be JSON only
- Scored by exact match after deterministic normalization

## 4. Task design principles

- Prefer tasks that require structure (blocks, hierarchy, adjacency)
- Avoid tasks that are solvable by surface text order alone
- Define canonical JSON outputs
- Use deterministic normalization for fairness

## 5. Scoring and normalization

- Normalize strings and JSON structure before comparison
- For unordered collections, compare as sorted sets
- Avoid ambiguous numbering in answers

## 6. Metrics

### 6.1 Primary metric: RUS

RUS = correct_answers / total_questions

### 6.2 Secondary metrics

- Cost-normalized RUS = RUS / cost_usd
- Token-normalized RUS = RUS / input_tokens
- Stage A failure rate = failed Markdown reconstruction rate

## 7. Directory layout

```
benchmark/
  rub/
    README.md
    BENCHMARK_SPEC.md
    manifest.json        # full (v1)
    manifest_lite.json   # lite (v0.1)
    truth/               # full (v1)
      *.json
    truth_lite/          # lite (v0.1)
      *.json
    schemas/
      *.schema.json
    scoring/
      normalize.py
      score.py
    diagrams/
      rub_overview.mmd
      scoring_flow.mmd
```

## 8. Manifest fields

- id: task id
- type: task type
- xlsx: input workbook path
- question: Stage B query
- truth: ground-truth JSON path
- sheet_scope: optional sheet filter (null = all)
- render: render settings for image/pdf paths
- track: evaluation track name (default: reconstruction)

## 8.1 RUB lite notes

- Smaller number of cases
- Unordered paths supported for strict but fair comparison
- Binary scoring (0/1) only

## 9. Evaluation notes

- Do not use Markdown string similarity for RUB scoring
- Focus on task correctness and structure preservation
- Keep normalization deterministic and transparent

## 10. Reporting

- Public report focuses on reconstruction utility
- Show both primary and secondary metrics
- Clearly separate core extraction vs RUB results
