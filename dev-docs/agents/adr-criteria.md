# ADR 判定基準

この文書は、AI エージェントが変更を `required` / `recommended` / `not-needed` に分類するための基準を定義する。

## 判定レベル

### `required`

次のいずれかに当てはまる場合は ADR 必須とする。

- 公開 API、CLI、MCP の契約や意味が変わる
- 出力 JSON / YAML / TOON の意味、必須性、省略方針が変わる
- 既定値が `default omit` / `default include` のように意味契約を変える
- mode の責務境界や validation 契約が変わる
- 同じ制約を API / CLI / engine / MCP の各入口でどう揃えるかが変わる
- backend の追加、削除、優先順位、fallback 方針が変わる
- runtime failure の扱い、理由コード、ログ整合、返却形状の組み合わせが変わる
- `PathPolicy` など safety boundary が変わる
- backward compatibility の方針が変わる
- patch/edit backend の capability や選択ポリシーが変わる

### `recommended`

設計判断の可能性は高いが、必ずしも ADR が必要とは限らない変更。

- テスト体系の大きな見直し
- エラー処理ポリシーの再整理
- 大きな依存ライブラリ追加や置き換え
- パフォーマンス最適化の基本戦略変更
- 運用フローや AI エージェント手順の恒久ルール化

### `not-needed`

policy-level change ではないため ADR 不要。

- typo 修正
- ふるまい不変の内部 refactor
- 既存契約を変えない文言整理
- 設計意図を変えない単純なテスト追加
- バグ修正のうち、既存 ADR / spec / tests の契約に戻すだけのもの

## exstruct 固有の必須領域

次の領域は、変更が入った時点で ADR 必須を第一候補として疑う。

| 領域 | 典型例 | 既存参照 |
| --- | --- | --- |
| extraction mode | `light` / `libreoffice` / `standard` / `verbose` の責務変更 | `ADR-0001`, `dev-docs/specs/excel-extraction.md` |
| backend fallback | COM / LibreOffice / 将来 backend の fallback 契約変更 | `ADR-0002` |
| serialization contract | metadata の省略/含有、schema の意味変更 | `ADR-0003`, `docs/api.md`, `docs/cli.md` |
| patch backend policy | `auto` / `com` / `openpyxl` の選択・制約変更 | `ADR-0004`, `docs/mcp.md` |
| safety boundary | path 許可範囲、正規化、出力位置のルール変更 | `ADR-0005`, `docs/mcp.md` |
| compatibility policy | backward compatibility の維持・破壊条件変更 | 既存 ADR または新規 ADR 対象 |

## 見分け方

次の問いに 1 つでも `yes` なら ADR を強く検討する。

1. これは「どう作るか」ではなく「なぜその方針を採るか」の変更か
2. 変更が `docs/specs/tests/src` の複数層へ波及するか
3. 将来、同じ論点で再び迷う可能性が高いか
4. 既存 ADR の前提や例外条件を壊していないか
5. レビューで「どの policy を守るべきか」が争点になりそうか

## 追加ヒューリスティクス

- mode や backend の変更が「どこで validation されるか」ではなく、「全入口で同じ validation を課すか」を変えるなら ADR 必須寄り
- runtime failure を扱う変更では、理由コード、ログ、返却データ形状の 3 点をセットで見る
- serialization の既定値変更は、小さく見えても consumer contract を変えるため ADR 必須寄り
- backend 選択の変更は、優先順位、禁止 fallback、事前 capability check のどれか 1 つでも触れたら ADR 必須寄り

## ADR 不要だが根拠は残すべきケース

- 既存 ADR の契約に戻すだけの回帰修正
- 既存 spec の誤記修正
- 既存 policy に従った tests 追加

この場合は、新規 ADR ではなく、関連する issue、PR、`tasks/todo.md`、レビューコメントに判断理由を短く残す。

## 判定出力の最小要件

AI エージェントは判定時に最低限次を返す。

- verdict: `required` / `recommended` / `not-needed`
- rationale: 1-3 行の理由
- affected domains: 関連する設計領域
- existing ADR candidates: 参照すべき既存 ADR
- next action: `new-adr`, `update-existing-adr`, `no-adr`
- evidence triad:
  - specs の契約文
  - src の主要シンボルまたは実行経路
  - tests の固定化ケース

## Phase 2: 監査 / 索引の出力要件

`adr-reconciler` は判定済み ADR の drift 監査で最低限次を返す。

- scope:
  - 対象 ADR
  - 調査した `specs` / `src` / `tests`
- findings:
  - type: `policy-drift` / `missing-adr-update` / `missing-evidence` / `stale-reference`
  - severity: `high` / `medium` / `low`
  - claim
  - affected ADRs
  - evidence matrix:
    - adr の claim または該当節
    - specs の契約文
    - src の主要シンボルまたは実行経路
    - tests の固定化ケース
  - recommended action: `update-adr` / `new-adr` / `update-specs` / `add-tests` / `no-action`

`adr-indexer` は索引更新で最低限次を返す。

- updated artifacts
- added or changed ADR entries
- consistency findings

## Phase 3: ドラフトレビューの出力要件

`adr-reviewer` は ADR 草案の設計レビューで最低限次を返す。

- prerequisite: 現行 draft に未解消の `adr-linter` `high` / `medium` finding が残っていない
- verdict: `ready` / `revise` / `escalate`
- scope:
  - 対象 ADR 草案
  - 調査した関連 ADR / `docs/` (公開 API / CLI / MCP が関係する場合) / `specs` / `src` / `tests`
  - 参照した issue / PR / diff context
- findings:
  - type: `decision-gap` / `scope-conflict` / `evidence-risk` / `rollout-gap` / `ownership-escalation`
  - severity: `high` / `medium` / `low`
  - summary
  - why it matters
  - suggested revision
  - evidence:
    - draft の該当節または claim
    - related sources
- open questions
- residual risks
