# ADR Review Contract

この文書は、`adr-reviewer` が ADR 草案をどうレビューし、どの形で findings を返すかを定義する。

## 目的

- `adr-linter` の構造検査と、ADR 草案の設計レビューを分離する
- review comment の観点を固定し、AI ごとに「何を問題とみなすか」がぶれないようにする
- AI が修正してよい論点と、人へ escalate すべき論点を明示する

## 非目標

- `adr-linter` の代わりに必須節や status の機械検査を再実装しない
- `adr-reconciler` のように merge 後 drift 監査をしない
- ADR 本文の自動修正を source of truth にしない

## Preconditions

`adr-reviewer` は、現行 draft が `adr-linter` で検査済みであり、未解消の `high` / `medium` finding が残っていない状態でのみ有効とする。

lint finding が残っている場合は、草案を修正して `adr-linter` を再実行してから設計レビューに進む。

## Review focus

`adr-reviewer` は次の観点を順に確認する。

1. decision coverage
   - ADR が本当に 1 つの policy decision を解決しているか
   - `why` と `how` を混同していないか
2. scope and lineage
   - 既存 ADR / spec と衝突、重複、説明不足がないか
   - supersede や update を使うべき論点を新規 ADR に分離し過ぎていないか
3. evidence strength
   - `Tests`, `Code`, `Related specs` が決定内容を本当に裏付けているか
   - consequence や claim が evidence なしに拡張されていないか
4. rollout and compatibility
    - 公開契約、fallback、migration、operational impact が relevant な場合に触れられているか
    - 公開 API / CLI / MCP が関係する場合は、対応する `docs/` を scope に含めて compatibility / break judgment の根拠を確認しているか
    - 非目標や移行しない理由が必要なのに省略されていないか
5. ownership boundary
    - AI の責務外である public API break judgement、security / license 判断、大規模ディレクトリ再編、未確定の product / spec 方針を勝手に確定していないか

## Verdict

`adr-reviewer` の verdict は次の 3 種類に固定する。

- `ready`
  - 高 / 中 severity の design finding がなく、AI の責務外論点も未解決でない
- `revise`
  - ADR 草案の内容を修正すれば解消できる review finding が残っている
- `escalate`
  - 人の判断が必要な論点が残っており、AI が draft だけでは閉じられない

`ready` は merge や最終承認の代替ではない。

## Result envelope

review 結果は少なくとも次の top-level fields を含む。

- `verdict`
- `scope`
- `findings`
- `open questions`
- `residual risks`

## Finding contract

findings は severity 順に返し、各 finding は次を含む。

- `type`
- `severity`
- `summary`
- `why it matters`
- `suggested revision`
- `evidence`
  - `draft`
  - `related sources`

finding type は次に固定する。

- `decision-gap`
  - ADR が解くべき policy question を十分に決めていない
- `scope-conflict`
  - 既存 ADR / spec と重複、衝突、責務混在がある
- `evidence-risk`
  - cited evidence が弱い、無い、または claim を支え切れていない
- `rollout-gap`
  - compatibility / migration / fallback / operational impact の説明が必要なのに欠けている
- `ownership-escalation`
  - AI の責務外論点を含み、人の判断が必要である

## Scope contract

review 結果には少なくとも次の scope を含める。

- 対象 ADR 草案
- 調査した関連 ADR
- 調査した関連 `docs/` (公開 API / CLI / MCP が関係する場合)
- 調査した `dev-docs/specs/`, `src/`, `tests/`
- 参照した issue / PR / diff context

## Relationship to other skills

- `adr-drafter` の後に使う
- `adr-linter` で現行 draft の未解消 `high` / `medium` finding を落とした後に使う
- 公開 surface が関係する場合は、関連する `docs/` を review scope に含める
- review 結果が policy-level drift を示す場合は `adr-reconciler` や `adr-suggester` に戻る
