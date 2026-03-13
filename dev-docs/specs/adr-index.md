# ADR 索引仕様

この文書は、`adr-indexer` が維持する derived artifact の内部契約を定義する。

## 対象 artifact

- `dev-docs/adr/README.md`
- `dev-docs/adr/index.yaml`
- `dev-docs/adr/decision-map.md`

source of truth は各 `ADR-xxxx-*.md` 本文であり、索引 artifact はそこから導出される。

## `index.yaml` 契約

- top-level key は `adrs` とする
- 各 entry は少なくとも次を持つ
  - `id`
  - `title`
  - `status`
  - `path`
  - `primary_domain`
  - `domains`
  - `supersedes`
  - `superseded_by`
  - `related_specs`
- `status` は `accepted`, `proposed`, `superseded`, `deprecated` のみを許可する
- `primary_domain` は `domains` のいずれか 1 つでなければならない
- `primary_domain` は README の「主ドメイン」列の source of truth とする
- `domains` は AI が related ADR をたどるための短い分類語に限定する
- `supersedes` と `superseded_by` は ADR ID の配列とし、片側だけに存在してはならない

## `decision-map.md` 契約

- domain ごとに ADR をグループ化する
- 見出しは `domains` 配列の各要素と 1 対 1 に対応し、複数 domain を 1 見出しへまとめない
- 各 ADR には少なくとも ID、タイトル、状態を載せる
- supersede 関係がある場合は、decision map 側にも明示する
- ADR が存在しない domain は無理に作らない

## `README.md` 契約

- 人間向けの短い一覧を持つ
- status と主ドメインは `index.yaml` の `status` / `primary_domain` と一致させる
- 詳細な relationship graph や機械可読 metadata は `decision-map.md` と `index.yaml` に委ねる

## 更新トリガー

次のいずれかが起きたら index artifact を更新する。

- 新規 ADR 追加
- ADR の status 変更
- `Supersedes` / `Superseded by` の変更
- ADR が参照する related spec の変更
- domain 分類の追加または見直し
