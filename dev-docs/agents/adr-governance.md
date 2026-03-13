# ADR ガバナンス

この文書は、AI エージェントが ExStruct で ADR をいつ作り、いつ更新し、何を根拠として扱うかを定義する。

## 目的

- 長期保守に影響する設計判断を、issue や差分の暗黙知に埋もれさせない
- `dev-docs/specs/`、`tests/`、`src/` との責務分離を維持する
- 既存 ADR と新しい変更の関係を明示し、判断理由の死文化を防ぐ

## 文書の役割

- ADR = なぜその方針を採用したか
- specs = 何を保証するか
- tests = その保証が実在する証拠
- src = どう実装するか

ADR に実装メモだけを書くのは禁止する。単なる手順や一時的な作業順序は `tasks/` に置く。

## ADR を新規作成する条件

次のいずれかに当てはまる場合は、新規 ADR を優先する。

- 新しい政策や責務境界を追加する
- 既存 ADR では説明できない別論点の判断を導入する
- 複数案の比較や長期トレードオフがある
- 将来も同じ論点が再発しうる

## 既存 ADR を更新する条件

次の条件を満たす場合は、既存 ADR の更新を優先する。

- 既存 ADR の政策は維持しつつ、背景、根拠、関連 spec を補強する
- 既存 ADR がカバーしている同じ論点に、追加 evidence や clarifying constraint を足す
- 実装や docs の整理により、参照先だけを最新化する

既存 ADR の結論そのものが変わる場合は、更新ではなく新規 ADR + supersede を検討する。

## Supersede の扱い

- 旧 ADR の判断がもはや権威を持たない場合は、新規 ADR を作成して旧 ADR を `superseded` にする
- 旧 ADR 側の `Superseded by` と、新 ADR 側の `Supersedes` を相互に更新する
- 部分上書きではなく、どの政策が置き換わったかを明記する

## Evidence 要件

ADR は少なくとも 1 つの concrete evidence を持つ。

- Tests: 回帰テスト、契約テスト、統合テスト
- Code: 判断を実装している主要ファイル
- Related specs: 現行仕様を定義している内部 spec または公開 docs

推奨は、`Tests`、`Code`、`Related specs` の 3 系統をすべて埋めること。

## Draft review ルール

`proposed` ADR を merge 前に扱うときは、構造検査と設計レビューを分ける。

- `adr-linter` は status、必須節、evidence の有無、supersede link の整合を検査する
- `adr-reviewer` は現行 draft に未解消の `adr-linter` `high` / `medium` finding がない状態で使い、公開 API / CLI / MCP に触れる ADR では関連 `docs/` も scope に含めて、判断の妥当性、既存 ADR / spec との衝突、evidence の説得力、互換性 / rollout / fallback / safety impact をレビューする
- `adr-reviewer` の verdict は `ready`, `revise`, `escalate` を使う
- `revise` の場合は、ADR 草案を更新して `adr-linter` と `adr-reviewer` を再実行する
- `escalate` は、AI の責務外である公開 API break judgement、security / license 判断、大規模ディレクトリ再編、未確定の product / spec 方針を含む場合に使う
- `ready` は merge 自体の承認を意味しない。人の最終判断が必要な論点は残りうる

## Reconciliation ルール

ADR が `proposed` または `accepted` の場合は、変更時または定期点検時に `adr-reconciler` で drift を確認する。

- 監査対象は claim 単位とし、`adr`, `specs`, `src`, `tests` の evidence matrix をそろえる
- finding 種別は少なくとも次を使う
  - `policy-drift`
  - `missing-adr-update`
  - `missing-evidence`
  - `stale-reference`
- findings は `severity` (`high` / `medium` / `low`) を持つ
- `high` findings は merge 前に解消するか、明示的な follow-up を作る
- 監査結果の `recommended action` は次を使う
  - `update-adr`
  - `new-adr`
  - `update-specs`
  - `add-tests`
  - `no-action`

`adr-reconciler` は監査結果を返すだけで、ADR や spec の本文を自動変更しない。policy-level change が疑われる場合は `adr-suggester` と `adr-drafter` に戻す。

## 索引 artifact の扱い

次のファイルは ADR 本文から導かれる derived artifact として扱う。

- `dev-docs/adr/README.md`
- `dev-docs/adr/index.yaml`
- `dev-docs/adr/decision-map.md`

これらは新規 ADR、status 変更、supersede 関係変更、domain 分類変更、related spec 更新があったときに更新する。

## Status ルール

利用可能な `状態` は次のみとする。

- `proposed`
- `accepted`
- `superseded`
- `deprecated`

`draft` や独自 status は使わない。草案段階でも ADR 文書上は `proposed` を使う。

## AI エージェントの必須動作

- 変更着手前に、関係しそうな既存 ADR を確認する
- ADR 対象の論点では、`why` と `what` を混同しない
- 新しい ADR を提案するときは、関連する `tests`, `code`, `specs` を列挙する
- ADR 不要と判断した場合も、なぜ policy-level change ではないのかを短く残す
- ADR を新規作成/更新した場合は、必要に応じて `adr-reconciler` で drift 監査を行い、`high` findings を残さない
- ADR を追加・更新・supersede した場合は、derived index artifact も同期させる
- ADR 草案をレビューするときは、現行 draft の `adr-linter` `high` / `medium` finding を先に解消し、公開 surface に触れる場合は関連 `docs/` も確認したうえで、`adr-reviewer` で設計論点だけを findings 化する

## 非目標

- ADR だけで公開契約を定義しない
- ADR だけで実装正当性を証明しない
- ADR を CI や bot 連携の詳細設計書として使わない
