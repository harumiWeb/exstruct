# ADR ワークフロー

この文書は、issue や PR から ADR を扱うときの標準フローを定義する。

## Phase 1 の対象

Phase 1 では次だけを標準化する。

1. ADR 要否判定
2. ADR 草案または既存 ADR 更新提案
3. ADR 文書の lint

## Phase 2 の追加対象

Phase 2 では、Phase 1 に次を追加する。

1. `adr-reconciler` による ADR と `specs` / `tests` / `src` の整合性監査
2. `adr-indexer` による ADR 索引と relationship map の更新

## Phase 3 の追加対象

Phase 3 では、Phase 1 + 2 に次を追加する。

1. `adr-reviewer` による ADR 草案の設計レビュー

## 標準フロー

1. issue または PR を読む
2. 関連する `docs/`, `dev-docs/specs/`, `dev-docs/adr/`, `tests/`, `src/` を読み、判定に必要な evidence triad を集める
3. `adr-suggester` で `required` / `recommended` / `not-needed` を判定する
4. `not-needed` の場合でも、判定理由と evidence triad を issue または PR に残す
5. `required` または `recommended` の場合は、`adr-drafter` で新規 ADR 草案または既存 ADR 更新提案を作る
6. `adr-linter` で形式と evidence を検査する
7. `adr-linter` に未解消の `high` / `medium` finding がある場合は、草案を更新して 6 を再実行する
8. `adr-reviewer` で設計妥当性、既存 ADR との衝突、互換性 / rollout / fallback / safety impact をレビューする。公開 API / CLI / MCP に触れる ADR では関連 `docs/` も scope に含める
9. `adr-reviewer` が `revise` を返した場合は、草案を更新して必要なら 6-8 を再実行する
10. `adr-reviewer` が `escalate` を返した場合は、人の判断が必要な論点として issue または PR に戻す
11. ADR が新規追加/更新された場合、または policy-level 変更を含む場合は `adr-reconciler` を実行する
12. merge 時に関連 spec / docs / tests と reconciliation findings の整合を再確認する
13. ADR が追加 / 更新 / supersede された場合は `adr-indexer` で `README.md`, `index.yaml`, `decision-map.md` を同期する

## 読み順

ADR 系タスクでは、次の順で資料を確認する。

1. `docs/`
2. `dev-docs/specs/`
3. `dev-docs/adr/`
4. `tests/`
5. `src/`

AI 向けの判断基準が必要なときだけ、追加で次を読む。

- `dev-docs/agents/adr-governance.md`
- `dev-docs/agents/adr-criteria.md`

## skill ごとの責務

### `adr-suggester`

- 変更を設計判断として扱うべきか判定する
- verdict 前に evidence triad を集める
- 新規 ADR 候補と既存 ADR 候補を返す
- `not-needed` を含め、判定結果には evidence triad を添える
- 草案本文は生成しない

### `adr-drafter`

- 新規 ADR 草案か既存 ADR 更新提案を作る
- `背景`, `決定`, `影響`, `根拠` を埋める
- 根拠には `Tests`, `Code`, `Related specs` を含める

### `adr-linter`

- `状態`、必須セクション、evidence、`Supersedes` / `Superseded by` を検査する
- 修正文案より findings を優先する

### `adr-reconciler`

- ADR の claim と `specs` / `src` / `tests` の現行状態を照合する
- finding ごとに `adr`, `specs`, `src`, `tests` の evidence matrix を返す
- finding 種別として `policy-drift`, `missing-adr-update`, `missing-evidence`, `stale-reference` を使う
- finding ごとに `severity` (`high` / `medium` / `low`) と `recommended action` を返す
- ADR 本文の自動修正は行わない

### `adr-reviewer`

- ADR 草案の設計レビューを行う
- 現行 draft に `adr-linter` の未解消 `high` / `medium` finding がないことを前提にする
- 公開 API / CLI / MCP に触れる ADR では関連 `docs/` を scope に含める
- findings 種別として `decision-gap`, `scope-conflict`, `evidence-risk`, `rollout-gap`, `ownership-escalation` を使う
- verdict は `ready`, `revise`, `escalate` を返す
- `adr-linter` が扱う構造検査を繰り返さず、設計論点、既存 ADR との衝突、互換性 / rollout / fallback / safety impact に集中する
- AI の責務外に触れる論点は `escalate` として人へ戻す

### `adr-indexer`

- 既存 ADR とその metadata を走査し、`README.md`, `index.yaml`, `decision-map.md` を同期する
- status、domain、supersede 関係、related specs の不整合を findings として返す
- 索引 artifact は source of truth ではなく、ADR 本文の derived view として扱う

## merge 前チェック

- ADR の結論が spec と矛盾していない
- spec に書かれた契約が tests で裏付けられている
- 既存 ADR を supersede する場合は相互参照が埋まっている
- ADR が不要な場合でも、その理由が issue または PR に残っている
- ADR が不要な場合でも、判定に使った `specs`, `src`, `tests` の根拠が追跡できる
- `adr-reconciler` の `high` findings が未解消のまま merge されていない
- `adr-linter` の `high` / `medium` findings が未解消のまま merge されていない
- `adr-reviewer` の `revise` verdict や `high` / `medium` findings が未解消のまま merge されていない
- `adr-reviewer` が `escalate` を返した論点が未処理のまま merge されていない

## merge 後 / 定期監査チェック

- `adr-reconciler` が `high` findings を返した場合は、対象 ADR と drift 元の spec / test / code path を issue または PR に残す
- drift が policy-level change を示す場合は、`adr-suggester` へ戻して `required` / `recommended` / `not-needed` を再判定する
- ADR が追加 / 更新 / supersede された場合は `adr-indexer` で derived artifact を更新する
- `index.yaml` と `decision-map.md` は、ADR 本文と異なる status や supersede 関係を持ってはならない

## 将来フェーズ

- review automation や PR bot 連携の詳細化
