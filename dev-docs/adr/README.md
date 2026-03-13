# ADR

ADR は、どの制約のもとで何を判断し、どのトレードオフを受け入れたかを記録する文書です。

## 目的

- 構造や長期保守に影響する判断を記録する
- 採用しなかった案と非目標を残す
- 将来どの ADR に置き換えられたかを追跡できるようにする

## 状態

- `proposed`
- `accepted`
- `superseded`
- `deprecated`

## 採番

- リポジトリ追加順に `ADR-0001`, `ADR-0002`, ... を使う
- タイトルが変わっても番号は固定する

## 他の文書との関係

- ADR は「なぜそうしたか」を説明する
- `dev-docs/specs/` は「何を保証するか」を説明する
- `tests/` はその振る舞いが実在する証拠を示す

## 索引 artifact

- `index.yaml`: AI が機械可読にたどるための ADR metadata index
- `decision-map.md`: domain ごとの関連 ADR と supersede 関係の概要

## 一覧

| ID | タイトル | 状態 | 主ドメイン |
| --- | --- | --- | --- |
| `ADR-0001` | 抽出モードの責務境界 | `accepted` | `extraction` |
| `ADR-0002` | Rich Backend のフォールバック方針 | `accepted` | `backend` |
| `ADR-0003` | 出力直列化における省略方針 | `accepted` | `schema` |
| `ADR-0004` | Patch backend 選択方針 | `accepted` | `mcp` |
| `ADR-0005` | PathPolicy の安全境界 | `accepted` | `safety` |
