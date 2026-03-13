# Todo

## 2026-03-13 issue #90 ADR management skills

### Planning

- [x] issue #90 の本文、ラベル、受け入れ条件を取得して要点を整理する
- [x] `dev-docs/adr/`、`dev-docs/agents/`、`.agents/skills/` の現状を確認し、既存土台と不足領域を把握する
- [x] Phase 1 の対象 skill と各 skill の責務境界を具体化する
- [x] exstruct 固有の ADR 必須 / 推奨 / 不要判定ルールを文書化する
- [x] `dev-docs/agents/` に追加する ADR 運用文書の構成と初期ドラフト方針を決める
- [x] Phase 1 の skill 雛形と参照導線を追加する
- [x] 文書整合性と skill frontmatter / metadata の最低限検証を行う

### Planning (Phase 2)

- [x] Phase 2 の対象を `adr-reconciler` と `adr-indexer` に確定する
- [x] `adr-reconciler` と `adr-indexer` の入出力契約を定義する
- [x] `.agents/skills/adr-reconciler/` と `.agents/skills/adr-indexer/` の `SKILL.md` / `agents/openai.yaml` を追加する
- [x] `dev-docs/agents/adr-workflow.md` と `adr-governance.md` に Phase 2 運用を反映する
- [x] `dev-docs/agents/adr-criteria.md` と `dev-docs/specs/adr-index.md` に Phase 2 の出力 / 索引契約を追加する
- [x] `dev-docs/adr/index.yaml` と `decision-map.md` の初期 artifact を追加する
- [x] issue #90 セクションの `tasks/feature_spec.md` / `tasks/todo.md` を更新し、検証結果を記録する

### Review follow-up

- [x] `adr-reconciler` の Output Contract に `scope` と `findings[].type` を追加する
- [x] `decision-map.md` の domain 見出しを `index.yaml.domains` と 1 対 1 にそろえる
- [x] `README.md` の主ドメインを決定する `primary_domain` ルールを `adr-index.md` と `index.yaml` に追加する
- [x] review 指摘反映後の validator / YAML parse / 差分チェックを再実行する

### Review

- issue #90 は、ADR 文書を 1 本生成する話ではなく、ADR 要否判定から草案作成、品質検査、整合性監査、索引化までを段階導入する skill 群の設計 issue として理解した。
- 既存土台:
  - `dev-docs/adr/` に ADR 本文、`README.md`、`template.md` が存在する
  - `dev-docs/agents/` に AI エージェント向け内部文書が存在する
  - `.agents/skills/` には既存 skill の配置例が 1 件ある
- 着手前のギャップ:
  - ADR ガバナンス専用 skill 群は未作成
  - `adr-governance.md`、`adr-criteria.md`、`adr-workflow.md` は未作成
  - ADR と実装 / tests / specs の整合性を継続監査する仕組みは未整備
- 実施した追加:
  - `dev-docs/agents/adr-governance.md`
    - ADR の目的、作成/更新条件、status ルール、evidence 要件、supersede 方針を定義した
  - `dev-docs/agents/adr-criteria.md`
    - `required` / `recommended` / `not-needed` の判定基準と、exstruct 固有の必須領域を定義した
  - `dev-docs/agents/adr-workflow.md`
    - Phase 1 の標準フローと skill ごとの責務を定義した
  - `.agents/skills/adr-suggester/`
    - issue / PR / diff の ADR 要否判定に特化した skill 雛形を追加した
  - `.agents/skills/adr-drafter/`
    - 新規 ADR 草案または既存 ADR 更新提案を作る skill 雛形を追加した
  - `.agents/skills/adr-linter/`
    - ADR 文書の構造・evidence・supersede 関係を検査する skill 雛形を追加した
  - `dev-docs/agents/README.md`
    - 新しい ADR 運用文書を参照順へ追加した
  - `tasks/feature_spec.md`
    - Phase 1 / Phase 2 以降の責務境界と追加成果物を明文化した
- 設計結論:
  - Phase 1 は `adr-suggester`、`adr-drafter`、`adr-linter` に限定し、Phase 2 で `adr-reconciler` と `adr-indexer`、Phase 3 で `adr-reviewer` を扱う形に整理した
  - exstruct 固有の必須領域は、mode 境界、backend fallback、serialization contract、patch backend policy、safety boundary、compatibility policy とした
  - 判定ヒューリスティクスとして、入口整合、fallback の理由コード/ログ/返却形状、serialization 既定値変更、backend 選択変更を ADR 必須寄りの信号として採用した
  - ADR 候補の裏付けは `specs` + `src` + `tests` の evidence triad で揃える方針にした
  - `adr-suggester` も `adr-drafter` と同様に evidence triad を成果物へ残し、`not-needed` 判定でも根拠追跡できる契約に修正した
  - `adr-workflow` は verdict 後ではなく verdict 前に evidence triad を集める流れへ修正した
  - Phase 2 として `adr-reconciler` と `adr-indexer` を追加し、継続監査と索引更新を標準フローに入れた
  - `adr-reconciler` の findings には `policy-drift`, `missing-adr-update`, `missing-evidence`, `stale-reference` の種別と、`severity` / `recommended action` を必須とする方針にした
  - `adr-indexer` は `README.md`, `index.yaml`, `decision-map.md` を同期する derived artifact manager と定義した
  - `dev-docs/specs/adr-index.md` に `index.yaml` / `decision-map.md` の内部契約を追加し、既存 ADR 5 本を seeded artifact として反映した
  - review follow-up として、`adr-reconciler` の Output Contract に `scope` と `findings[].type` を追加し、`adr-indexer` の machine-readable metadata に `primary_domain` を追加した
  - `decision-map.md` は `domains` 配列の各要素を独立見出しとして表現し、`README.md` の主ドメインは `index.yaml.primary_domain` を source of truth にする方針へ修正した
- 検証:
  - `.agents/skills/` のディレクトリ一覧で `adr-reconciler` と `adr-indexer` が追加されていることを確認した
  - `python C:\Users\HARUMI\.codex\skills\.system\skill-creator\scripts\quick_validate.py .agents\skills\adr-reconciler`
  - `python C:\Users\HARUMI\.codex\skills\.system\skill-creator\scripts\quick_validate.py .agents\skills\adr-indexer`
  - `python` で `dev-docs/adr/index.yaml` と `agents/openai.yaml` 2 本の YAML parse が通ることを確認した
  - `git diff --check -- <changed files>` で実害のある差分エラーがないことを確認した
  - `tasks/feature_spec.md`, `dev-docs/agents/`, `dev-docs/specs/adr-index.md`, `.agents/skills/` の参照更新を目視確認した
  - `adr-reconciler` と `adr-indexer` の契約用語が `SKILL.md`、`adr-workflow.md`、`adr-governance.md`、`adr-criteria.md`、`adr-index.md` で一致していることを確認した

### Future phase backlog

- [ ] Phase 3 で `adr-reviewer` の入出力契約を定義し、ドラフトレビュー観点を固定する

## 2026-03-13 PR #91 unresolved review follow-up

### Planning

- [x] PR #91 の未 resolve review thread を最新状態で取得し、論点を列挙する
- [x] 各 thread を現行コードとテストで照合し、採用/非採用を分類する
- [x] 採用した指摘だけを最小差分で修正し、必要な検証を行う
- [x] 非採用の thread には判断理由を返信して resolve する
- [x] 実施内容、検証結果、未解決事項を Review セクションに記録する

### Review

- PR #91 の review thread は、初回 16 件に加えて push 後の CodeRabbit follow-up 2 件を含め、最終的に 18 件すべてを resolve した。
- 初回 16 件の分類は、13 件を採用、2 件を重複、1 件を対応不要と判定した。
- 採用した修正:
  - `dev-docs/specs/excel-extraction.md`
    - `libreoffice` mode を追加
    - rich backend failure 時に保持される pre-com artifact を実装どおり明記
    - `light` mode の説明を pre-com artifact を含む形へ修正
  - `dev-docs/agents/contributing.md`
    - `integrate.py` と `modeling.py` の責務を実装どおりに修正
  - `dev-docs/agents/overview.md`
    - mode 一覧と backend 説明に `libreoffice` を追加
  - `dev-docs/architecture/contributor-guide.md`
    - backend ファイルパスを `core/backends/*` に修正
  - `dev-docs/architecture/overview.md`
    - アーキテクチャ概要に LibreOffice backend を追加
  - `dev-docs/testing/test-requirements.md`
    - coverage target を `80%` に修正
    - `CellRow.r` を 1-based に修正
    - pandas dtype を `dtype=str` に修正
    - merged cell の空値正規化を `" "` に修正
    - 参照 spec 名を `dev-docs/specs/data-model.md` に修正
    - Windows smoke contract の `soffice.com` 優先 / `soffice.exe` fallback を現行運用に合わせて修正
  - `dev-docs/specs/patch/model-migration-notes.md`
    - canonical model が `patch/models.py` に既にあり、`internal.py` に重複定義が残っている現状へ修正
    - 移行手順を「移設」ではなく「重複解消と import 統一」へ修正
- 非採用/重複:
  - `dev-docs/agents/overview.md` の CodeRabbit mode 指摘は Copilot 指摘と重複
  - `dev-docs/specs/excel-extraction.md` の CodeRabbit mode 指摘は Copilot 指摘と重複
  - `tasks/todo.md` 上の PR Summary 指摘は、現在の PR 本文がすでに `dev-docs/` 配下の internal docs 追加を明記しているため対応不要
- 検証:
  - `gh api graphql` で PR #91 の unresolved review thread を取得
  - 現行コード照合:
    - `codecov.yml`
    - `src/exstruct/core/cells.py`
    - `src/exstruct/core/integrate.py`
    - `src/exstruct/core/modeling.py`
    - `src/exstruct/core/pipeline.py`
    - `src/exstruct/mcp/patch/models.py`
    - `src/exstruct/mcp/patch/internal.py`
    - `src/exstruct/mcp/patch_runner.py`
    - `src/exstruct/mcp/patch/service.py`
  - `git diff --check -- <changed files>` で whitespace error なしを確認
- GitHub:
  - commit `1a8e6ba` と `3bad843` を `origin/docs/adr` へ push 済み
  - 最終確認で review thread 18 件すべて `isResolved: true`
  - 対応不要の PR Summary 指摘は thread reply 権限がなかったため、PR comment `#issuecomment-4052594372` で理由を補足した
- 追加 follow-up:
  - push 後に CodeRabbit から `test-requirements.md` へ 2 件の新規 thread が追加されたが、`dev-docs/specs/data-model.md` 参照名と Windows smoke の `soffice.com` 優先契約へ追随して解消した

## 2026-03-13 AI agent document retention policy

### Planning

- [x] 既存の `AGENTS.md`、`tasks/feature_spec.md`、`tasks/todo.md` を確認し、現状の task 文書が担っている役割を整理する
- [x] 残すべき設計判断と破棄可能な作業メモの境界を定義する
- [x] `AGENTS.md` に、文書保持ポリシー、ADR 化条件、白紙化禁止の粒度を追記する
- [x] 変更内容を確認し、今回の運用変更がそのまま AI 指示として使えることを確認する

### Review

- `AGENTS.md` に `## ドキュメント保持ポリシー` を追加し、`tasks/feature_spec.md` と `tasks/todo.md` をファイル単位で白紙化しない運用を明文化した。
- task 文書の役割分離、残すべき情報、捨ててよい情報、完了時の移管手順、ADR 作成条件、セッション終了チェックを一か所にまとめた。
- `dev-docs/README.md` の整理に合わせ、内部向け恒久文書の移管先を `dev-docs/adr/`、`dev-docs/specs/`、`dev-docs/architecture/` に揃え、公開契約だけを `docs/*.md` に残す方針へ修正した。
- 検証:
  - `AGENTS.md` の追記内容を目視確認し、役割分離、残す/捨てる基準、完了時の必須手順、ADR 条件が含まれていることを確認した
  - doc-only 変更のためテストは未実行

## 2026-03-12 v0.6.1 release notes and changelog

### Planning

- [x] `v0.6.0..314ac86` の差分を確認し、`v0.6.1` で告知する変更点を抽出する
- [x] `tasks/feature_spec.md` に今回のリリース文書方針を記録する
- [x] `CHANGELOG.md` に `v0.6.1` エントリを追加する
- [x] `docs/release-notes/v0.6.1.md` を作成し、patch release の要点をまとめる
- [x] `mkdocs.yml` の release notes nav を更新する
- [x] 変更差分を確認し、最低限の文書整合性検証を行う

### Review

- `CHANGELOG.md` には `2026-03-12` 付けで `0.6.1` セクションを追加し、PR #79 系列の差分を `Added / Changed / Fixed` に整理した。
- `docs/release-notes/v0.6.1.md` を新規追加し、Windows LibreOffice smoke job 追加、runtime/path discovery 安定化、bridge subprocess の Windows 対応、smoke gate false-negative 抑制を patch release の要点として記載した。
- `mkdocs.yml` の Release Notes nav に `v0.6.1` を追加し、公開サイトから新しい release notes に遷移できるようにした。
- 検証:
  - `python -c "import yaml, pathlib; yaml.safe_load(pathlib.Path('mkdocs.yml').read_text(encoding='utf-8')); print('yaml-ok')"` -> `yaml-ok`
  - `git diff --check` -> 空白エラーなし。LF/CRLF の警告のみ
  - 目視確認で `CHANGELOG.md` 冒頭、`docs/release-notes/v0.6.1.md`、`mkdocs.yml` の nav 追加を確認した

## 2026-03-10 PR #76 review + Codacy follow-up

### Planning

- [x] review thread 2 件と Codacy `Bandit_B404` notice の妥当性を現行コードで再確認する
- [x] `libreoffice_backend.py` の partial order fallback と direction fallback を修正する
- [x] `tests/core/test_libreoffice_backend.py` に回帰テストを追加/更新する
- [x] `src/exstruct/core/libreoffice.py` の `subprocess` import notice を narrow suppression で解消する
- [x] 対象 pytest と `uv run task precommit-run` で検証する

### Review

- GitHub の未 resolve review thread 2 件はどちらも採用した。
  - `_match_by_name_then_order()` は name match 後の残差に対して、件数不一致でも `zip(..., strict=False)` で partial order fallback を適用するようにした。
  - connector `direction` は UNO bbox の `width/height` からは推定せず、OOXML delta が無いときは resolved endpoint shape centers から導く形へ変更した。endpoint geometry まで無い場合は `None` を返す。
- `src/exstruct/core/libreoffice.py` の `import subprocess` に `# nosec B404` コメントを追加し、Codacy/Bandit の import-level notice を narrow suppression した。
- `tests/core/test_libreoffice_backend.py` では次を回帰固定した。
  - partial order fallback が件数不一致でも動くこと
  - zero-length OOXML delta が resolved endpoint geometry に fallback すること
  - direction metadata が取れないときに bbox から誤推定せず `None` を返すこと
  - resolved endpoint geometry だけでも direction が計算できること
- 検証:
  - `python scripts/codacy_issues.py --pr 76 --min-level Info`（修正前の remote PR 状態）では `Bandit_B404` 1 件を確認
  - `wsl.exe bash -lc 'cd /mnt/c/dev/Python/exstruct && uv run pytest tests/core/test_libreoffice_backend.py -q -k "match_by_name_then_order or ooxml_zero_delta_direction_falls_back_to_resolved_shape_geometry or resolve_direction_returns_none_without_ooxml_delta_or_resolved_shapes or resolve_direction_uses_resolved_shape_geometry_without_ooxml_metadata or combines_ooxml_and_uno_connector_endpoints or rotates_ooxml_connector_delta_for_heuristic_matching or resolve_direction_uses_unrotated_ooxml_delta or resolve_direction_rotates_ooxml_delta_before_mapping" --basetemp .pytest-tmp-codex-review'` -> `8 passed, 39 deselected`
  - `wsl.exe bash -lc 'cd /mnt/c/dev/Python/exstruct && uv run task precommit-run'` -> `ruff / ruff-format / mypy passed`
- 補足:
  - Codacy の remote PR status は未 push のためまだ再評価されていない。今回の import-level suppression が反映されるのは次回 push 後。

## 2026-03-09 LibreOffice stderr cleanup masking fix

### Planning

- [x] cleanup maskingの契約を `tasks/feature_spec.md` に追記する
- [x] stderr log unlink を best-effort retry に変更する
- [x] cleanup lock が startup failure を隠さない回帰テストを追加する
- [x] 影響範囲の手動検証を行う

### Review

- `src/exstruct/core/libreoffice.py` に stderr log unlink 用の短い retry budget を追加し、`PermissionError` は cleanup failure として握りつぶすようにした
- `_close_stderr_sink()` は `stderr_path.unlink()` を直接呼ばず、best-effort helper 経由で削除するようにした
- `tests/core/test_libreoffice_backend.py` に以下の regression tests を追加した
  - 一時的な `PermissionError` 後に unlink 成功へ回復する test
  - unlink が常に lock されても `_start_soffice_startup_attempt()` が `PermissionError` ではなく元の startup failure を返す test
- 検証
  - `python -m py_compile src/exstruct/core/libreoffice.py tests/core/test_libreoffice_backend.py`
  - 対象テスト関数の直接実行
    - 新規 2 tests: pass
    - 既存回帰相当 2 tests (`test_libreoffice_session_enters_with_isolated_profile`, `test_probe_uno_bridge_handshake_uses_bridge_script`): pass
  - 備考: `pytest` 本体はこの環境で tmpdir / basetemp cleanup が `PermissionError` になり、通常実行では harness 側が不安定だった

## 2026-03-09 LibreOffice WinError32 investigation

### Planning

- [x] `libreoffice` モード失敗をローカルで再現する
- [x] `src/exstruct/core/libreoffice.py` の startup / cleanup 経路を確認する
- [x] `PermissionError(13)` の直接発生箇所と一次原因を切り分ける
- [x] 阻害要因を `tasks/todo.md` に記録する

### Review

- 再現は `sample/basic/sample.xlsx` を絶対パスで `mode="libreoffice"` 実行したときに確認した
- 直接の `PermissionError(13)` は `src/exstruct/core/libreoffice.py` の `_close_stderr_sink()` にある `stderr_path.unlink()` で発生していた
- ただし一次原因は cleanup 失敗ではなく、その直前の LibreOffice startup failure だった
- `soffice` の stderr には、isolated profile 起動時に `UserInstallation` 用ディレクトリへアクセスできず起動できない旨の Fatal Error が出ていた
- `_close_stderr_sink()` の `PermissionError` が上位に伝播し、本来の `LibreOfficeUnavailableError` を覆い隠して `libreoffice_pipeline_failed` に化けていた
- shared profile fallback もこの環境では `soffice exited during startup` で失敗した

## 2026-03-08 LibreOffice subprocess Codacy false-positive follow-up

### Planning

- [x] `libreoffice.py` の `subprocess.run(...)` call site と関連テストを確認する
- [x] `tasks/feature_spec.md` に helper 分解方針を追記する
- [x] 汎用 `_run_trusted_subprocess(...)` を bridge/probe 専用 helper へ分解する
- [x] 既存テストを新 helper 構造に合わせて更新する
- [x] 対象 pytest と `uv run task precommit-run` で検証する

### Review

- `src/exstruct/core/libreoffice.py` の汎用 `_run_trusted_subprocess(args, ...)` を廃止し、以下の専用 helper へ分解した
  - `_run_soffice_version_subprocess(...)`
  - `_run_bridge_probe_subprocess(...)`
  - `_run_bridge_extract_subprocess(...)`
  - `_run_bridge_handshake_subprocess(...)`
- bridge extraction は workbook path を `--file` argv で渡すのをやめ、`--file-stdin` + stdin テキストで bridge に渡す形へ変更した
- `src/exstruct/core/_libreoffice_bridge.py` は `--file-stdin` を受け付け、stdin から workbook path を読むようにした
- 追加・更新した検証
  - `tests/core/test_libreoffice_backend.py` で extraction helper の固定 argv / allowlisted env / stdin input を確認
  - `tests/core/test_libreoffice_bridge.py` で `--file-stdin` の parse と `main()` の stdin 読み取りを確認
  - 既存の bridge extraction / invalid JSON / handshake / probe 系テストは helper 分解後も通ることを確認
- ローカル検証結果
  - `uv run pytest tests/core/test_libreoffice_bridge.py tests/core/test_libreoffice_backend.py -q` -> `49 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
- Codacy 再取得 (`python scripts/codacy_issues.py --pr 76 --min-level Warning`) は remote PR 状態を返すため、ローカル未 push 状態では旧 issue 1 件が残る

## 2026-03-08 PR #76 unresolved thread follow-up

### Planning

- [x] PR #76 の未 resolve review thread を再取得して現行コードと突き合わせる
- [x] `tasks/feature_spec.md` に今回の follow-up contract を追記する
- [x] connector heuristic endpoint matching に rotation 反映を追加する
- [x] LibreOffice pipeline で rich artifact の部分成功を保持する
- [x] connector / pipeline fallback の regression tests を追加する
- [x] 関連 pytest と `uv run task precommit-run` で検証する

### Review

- PR #76 の未 resolve thread は 2026-03-08 時点で 2 件だった
  - `src/exstruct/core/backends/libreoffice_backend.py:539-541`
  - `src/exstruct/core/pipeline.py:1009-1010`
- 実装結果
  - `src/exstruct/core/backends/libreoffice_backend.py` は OOXML connector heuristic endpoint の `dx/dy` に `rotation` を反映してから begin/end 候補点を計算するようにした
  - `src/exstruct/core/pipeline.py` は LibreOffice rich backend 解決後の `extract_shapes()` と `extract_charts()` を分離し、chart failure 時は取得済み shape artifact を保持したまま fallback workbook を組み立てるようにした
  - `tests/core/test_libreoffice_backend.py` に回転付き heuristic endpoint regression test を追加した
  - `tests/core/test_pipeline_fallbacks.py` に `shapes success + charts failure` と `shapes failure short-circuit` の regression tests を追加した
- 検証結果
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py -q` -> `45 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`

## 2026-03-08 coverage / Codacy / CodeRabbit follow-up triage

### Planning

- [x] 最新 GitHub Actions run `22815422942` の failure reason を確認する
- [x] local non-COM coverage report を再実行し、coverage 低下の主戦場を特定する
- [x] PR #76 の最新 Codacy 指摘を再取得し、runtime defect と static-analysis false positive を切り分ける
- [x] PR #76 の最新 CodeRabbit / review thread を再確認し、未解決 thread と duplicate review comment を切り分ける
- [x] `tasks/feature_spec.md` に今回の accepted follow-up / out-of-scope を追記する
- [x] `engine.py` の per-call override 伝播を shared state mutation なしに組み直す
- [x] `engine.py` の immutability regression tests を追加する
- [x] LibreOffice startup 後の UNO bridge handshake を追加し、wrong-listener false positive を retry failure に戻す
- [x] wrong-listener / handshake failure の regression tests を追加する
- [x] `src/exstruct/core/libreoffice.py` の Codacy 対象 4 call site に最小スコープ suppression/comment を入れる
- [x] coverage 回復用の targeted tests を `libreoffice.py` / `libreoffice_backend.py` / `engine.py` に追加する
- [x] local non-COM suite coverage が `>= 80%` に戻ることを確認する

### Review

- 最新 GitHub Actions run は `2026-03-08` の `pytest` run `22815422942`
  - `test (ubuntu-latest, 3.12)` の `Run tests (non-COM suite)` が失敗
  - test 自体は `783 passed, 3 skipped, 11 deselected`
  - failure reason は `Coverage failure: total of 79 is less than fail-under=80`
  - Actions log の最終値は `Total coverage: 78.80%`
  - `ubuntu-latest, 3.11` / `windows-latest` matrix は fail-fast で `cancelled`
- ローカル再現:
  - `uv run pytest -m "not com and not render" --maxfail=1 --disable-warnings -q --cov=exstruct --cov-report=term-missing:skip-covered --cov-fail-under=0`
  - `785 passed, 1 skipped, 11 deselected`
  - coverage の低い変更対象は `src/exstruct/core/libreoffice.py (67%)`, `src/exstruct/core/backends/libreoffice_backend.py (79%)`, `src/exstruct/core/ooxml_drawing.py (83%)`
- Codacy (`python scripts/codacy_issues.py --pr 76 --min-level Warning`) は 7 件
  - `Bandit_B603` warning: `src/exstruct/core/libreoffice.py:267`, `:348`, `:398`
  - Semgrep error: `src/exstruct/core/libreoffice.py:267`, `:268`, `:348`, `:398`
  - 現行コードは `shell=False` + argv list + allowlisted env であり、今回は runtime exploit より analyzer 側の trust-boundary 誤検知として扱う
- CodeRabbit / review で現在も効いている追加指摘は 2 点
  - unresolved thread `src/exstruct/engine.py:250-265` (`also applies to 425-442`)
    - `_process_extract_scope()` が `self.output.destinations.auto_page_breaks_dir` を mutate し、immutable contract と caller isolation を壊す
  - latest review duplicate comment `src/exstruct/core/libreoffice.py:391-413` (`also applies to 768-835`)
    - `_reserve_tcp_port()` 後に別 listener が先に bind した場合でも `_wait_for_socket()` が startup success と誤認しうる
- 今回の方針
  - coverage gate は threshold 緩和ではなく targeted tests で戻す
  - `process()` の override 伝播は engine-level seam を維持しつつ explicit parameter 化し、shared state mutation を除去する
  - startup false positive は PID lookup ではなく UNO bridge handshake で詰める
  - Codacy は helper を前提に、該当 call site のみ inline suppression/comment で収束させる
- 実装結果
  - `src/exstruct/engine.py` は `_process_extract_scope()` を削除し、per-call destination がある場合だけ private override を `extract()` に渡す形へ変更した
  - `src/exstruct/core/libreoffice.py` は trusted subprocess helper と `--handshake` ベースの UNO startup 検証を追加し、wrong-listener を retry failure として扱う
  - `src/exstruct/core/_libreoffice_bridge.py` は `--handshake` / `--connect-timeout` を追加し、source file 自体の unit tests を追加した
- 検証結果
  - `uv run pytest tests/core/test_libreoffice_bridge.py tests/backends/test_auto_page_breaks.py tests/core/test_libreoffice_backend.py tests/engine/test_engine.py -q` -> `59 passed`
  - `uv run task precommit-run` -> `ruff`, `ruff-format`, `mypy` passed
  - `uv run pytest -m "not com and not render" --maxfail=1 --disable-warnings -q --cov=exstruct --cov-report=term-missing:skip-covered --cov-fail-under=80` -> `797 passed, 1 skipped, 11 deselected`, `Total coverage: 80.29%`

## 2026-03-08 PR #76 post-push triage

### Planning

- [x] 最新 GitHub Actions run / Codacy / review thread を再収集する
- [x] blocking issue と minor nit を切り分ける
- [x] `tasks/feature_spec.md` に post-push follow-up contract を追記する
- [x] `ExStructEngine.process(...)` が engine-level `extract(...)` seam を bypass しないように修正する
- [x] `tests/engine/test_engine.py::test_engine_process_normalizes_string_paths` の回帰を再現・固定する
- [x] `_read_relationships(...)` を typed relationship へ変更し、call site を relationship type 判定に切り替える
- [x] `tests/conftest.py` の LibreOffice runtime probe で broad `except Exception` をやめ、unexpected regression を surfacing する
- [x] `docs/api.md` の CLI 対応例に `--include-backend-metadata` を反映する
- [x] `_merge_anchor_geometry(...)` を anchor-first placement に揃え、位置マッチング回帰 test を追加する
- [x] `src/exstruct/core/libreoffice.py` の subprocess 呼び出しを trust-boundary 明示付き helper / suppression 方針で整理し、Codacy blocking issue を解消する

### Review

- 最新 GitHub Actions run は `2026-03-08` の `pytest` run `22814113410`
  - `test (ubuntu-latest, 3.12)` が `Run tests (non-COM suite)` で失敗
  - 実失敗は `tests/engine/test_engine.py::test_engine_process_normalizes_string_paths`
  - 例外は `ValueError: Excel file format cannot be determined, you must specify an engine manually.`
  - `windows-latest` jobs は fail-fast により `cancelled`
- Codacy は PR #76 に対して 4 件の `Error` を返している
  - 対象はすべて `src/exstruct/core/libreoffice.py`
  - ルールは `dangerous-subprocess-use-audit` と `dangerous-subprocess-use-tainted-env-args`
  - `shell=True` や command injection は現状見当たらないため、根本論点は subprocess の trust-boundary を code 上で明示できていない点と判断する
- 追加 review で妥当と判断した項目
  - `engine.process(...)` が engine-level `extract(...)` seam を bypass して test contract を壊している
  - `_read_relationships(...)` が `Type` を捨て、filename 推測に依存している
  - `tests/conftest.py` の broad `except Exception` が probe regression を skip に潰しうる
  - `docs/api.md` の CLI 例が `include_backend_metadata=True` と一致していない
  - `_merge_anchor_geometry(...)` が left/top を child-first のままにしている
- 今回 scope 外に置く項目
  - `pytest.MonkeyPatch` への import cleanup
  - `_start_soffice_startup_attempt(...)` の readability-only helper 分割
- 2026-03-08 実装・検証
  - `process()` は validation 後に engine-level `extract(...)` seam を通すよう修正し、string path 回帰 test を維持した
  - `_read_relationships(...)` は `Type` を保持する typed metadata へ変更し、sheet / drawing / chart 解決を relationship type 判定に切り替えた
  - `tests/conftest.py` は expected な unavailable failure だけを skip 扱いにし、unexpected probe regression は例外として surface させた
  - `docs/api.md` の CLI 相当例に `--include-backend-metadata` を反映した
  - `_merge_anchor_geometry(...)` は left/top を anchor-first に寄せ、回帰 test を追加した
  - LibreOffice subprocess は allowlist 環境変数と正規化済み executable/path helper を経由させ、bridge probe/env の回帰 test を追加した
  - `uv run pytest tests/engine/test_engine.py tests/test_conftest_libreoffice_runtime.py tests/core/test_libreoffice_backend.py -q` -> `44 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`

## 2026-03-08 PR #76 triage

### Planning

- [x] PR #76 の CI / Codacy / review 指摘を収集する
- [x] 各指摘を `feat/libreoffice-mode` の現行コードと GitHub ruleset に照らして妥当性判定する
- [x] follow-up に採用する項目を `tasks/feature_spec.md` に仕様として整理する
- [ ] `ExStructEngine.process(...)` と `extract(...)` の auto page-break override 不整合を修正する
- [ ] LibreOffice guardrail の test を拡張し、`image` / `.xls` / passthrough / extraction validator 分岐を直接カバーする
- [ ] `docs/cli.md` / `docs/api.md` の public surface 差分を解消する
- [ ] `confidence` の model constraint を追加し、schema 生成物と関連 test を同期する
- [ ] OOXML parsing hardening (`defusedxml`, anchor placement, scatter/bubble `xVal/yVal`) を実装する
- [ ] connector endpoint / direction の edge case を修正する
- [ ] LibreOffice session の `stderr` handling と smoke skip gate を harden する
- [ ] GitHub ruleset に `libreoffice-linux-smoke` を required status check として追加する
- [ ] OOXML-only shape/connector merge 方針と port reservation race の扱いを追加調査する

### Review

- 2026-03-08 時点の CI failure は `test (ubuntu-latest, 3.12)` の coverage 不足 (`78.62% < 80%`) であり、pytest 自体は `763 passed, 3 skipped, 11 deselected` だった
- Codacy は PR #76 に対して 1 件のみ `Error` を返し、対象は `src/exstruct/core/ooxml_drawing.py` の XML parser hardening だった
- 妥当と判断した主な指摘:
  - `process(..., auto_page_breaks_dir=...)` の override が extraction に伝播していない
  - guardrail test の分岐不足
  - `docs/cli.md` / `docs/api.md` の公開ドキュメントずれ
  - `libreoffice-linux-smoke` の skip 許容と ruleset 未設定
  - connector 片側未解決時の早期 return
  - `(dx, dy) = (0, 0)` の方向推定
  - `stderr=subprocess.PIPE` 未ドレイン
  - OOXML anchor placement 未使用
  - scatter/bubble `xVal` / `yVal` 未対応
  - `confidence` 範囲未制約
  - `xml.etree.ElementTree` の hardening 未対応
- 今回 scope 外とした主な指摘:
  - `ShapeData` / `ChartData` の dataclass 化
  - `load_workbook()` / `close_workbook()` typed handle 化
  - `normalize_path()` docstring のみの修正
  - `.xls` 例外型の `ConfigError` への変更

### 2026-03-08 Implementation Progress

- [x] `process()` / `extract()` の auto page-break override 整合を修正
- [x] LibreOffice guardrail tests (`image` / `.xls` / passthrough / extraction validator / process propagation) を追加
- [x] `docs/cli.md` / `docs/api.md` の follow-up docs 差分を反映
- [x] `confidence` の model/schema range 制約と validation/schema tests を追加
- [x] OOXML parser hardening (`defusedxml`) と anchor / scatter series fallback を実装
- [x] connector endpoint の mixed direct resolve と `(0,0)` direction fallback を修正
- [x] LibreOffice `stderr` を temp sink 化して undrained pipe を解消
- [x] LibreOffice smoke job を fail-fast 化し、GitHub ruleset に `libreoffice-linux-smoke` required status check を追加
- [ ] OOXML-only append 方針と `_reserve_tcp_port()` race は調査継続

### 2026-03-08 Verification Addendum

- `uv run pytest tests/test_constraints.py tests/backends/test_auto_page_breaks.py tests/core/test_libreoffice_backend.py tests/models/test_models_validation.py tests/models/test_schemas_generated.py tests/test_conftest_libreoffice_runtime.py tests/core/test_mode_output.py -k "libreoffice or process_excel_rejects or extract_workbook_rejects or not test_standard and not test_verbose" -q` -> `59 passed, 2 deselected`
- `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
- `gh api repos/harumiWeb/exstruct/rulesets/11087410` で `required_status_checks.context=libreoffice-linux-smoke` を確認

### 2026-03-08 Follow-up Planning

- [x] `_reserve_tcp_port()` race hardening を実装する
- [x] startup attempt 内の bounded retry contract を code / spec / tests に反映する
- [x] port collision を模した retry success / retry exhausted regression tests を追加する
- [x] draw-page snapshot あり sheet で unmatched OOXML shape / connector を append しない contract test を追加する
- [x] unmatched OOXML 件数の debug logging を追加する

### 2026-03-08 Follow-up Review

- `_reserve_tcp_port()` は hold-open reservation ではなく retry-based hardening で扱う
- OOXML-only append は v1 非対応とし、UNO canonical order を崩さない contract を固定する
- append を見送る代わりに、必要なら unmatched 件数の観測だけ別タスクで扱う
- `src/exstruct/core/backends/libreoffice_backend.py` で snapshot あり sheet の unmatched OOXML shape / connector 件数を `debug` logging するようにし、emit contract は変えずに観測だけ追加した
- 2026-03-08 follow-up implementation:
  - `src/exstruct/core/libreoffice.py` に startup strategy 内の bounded retry (`3` 回、short backoff) を追加し、失敗のたびに新しい port で `soffice` を起動し直すようにした
  - retry failure は `attempt N/3` 形式で集約し、strategy 間の failure aggregation と合わせて最終エラー文言に残すようにした
  - `tests/core/test_libreoffice_backend.py` に retry-within-strategy / retry-exhausted-to-shared-profile / aggregated-failure-detail の regression test を追加した
  - snapshot あり sheet では OOXML-only shape / connector を append しない contract test を追加し、v1 の emit 順序を固定した

## 2026-03-06 Backend Metadata Output Follow-up

### Planning

- [x] shape/chart backend metadata の常時出力箇所を調査する
- [x] `include_backend_metadata` を Python API / engine / CLI / MCP の共通オプションとして追加する
- [x] 既定値を metadata 非表示に切り替え、明示指定時のみ出力する
- [x] 直列化 helper / model helper / per-sheet / per-area export の挙動を揃える
- [x] README / API docs / MCP docs / release note / spec を更新する
- [x] 関連 unit test を追加・更新する

### Review

- `serialize_workbook`, `save_sheets`, `save_print_area_views`, `save_auto_page_break_views` に `include_backend_metadata` を追加し、shape/chart metadata を出力時だけ制御するようにした
- `SheetData` / `WorkbookData` / `PrintAreaView` の `to_json` / `to_yaml` / `to_toon` / `save` に同フラグを通し、raw model 直列化の既定値も揃えた
- `FilterOptions.include_backend_metadata`, CLI `--include-backend-metadata`, MCP `options.include_backend_metadata` を追加した
- 追加検証:
  - `uv run pytest tests/models/test_models_export.py tests/engine/test_engine.py tests/cli/test_cli.py tests/mcp/test_extract_runner_utils.py tests/mcp/test_tool_models.py tests/core/test_mode_output.py -k "not mark.com and not standard and not verbose and not line and not connector" -q`
  - `uv run pytest tests/io/test_print_area_views.py tests/core/test_error_handling_exceptions.py tests/export/test_export_requirements.py tests/mcp/test_tools_handlers.py tests/mcp/test_extract_alpha_col.py -q`

## 2026-03-06 README Sync

### Planning

- [x] `README.ja.md` を基準に `README.md` / `docs/README.ja.md` / `docs/README.en.md` の差分を整理する
- [x] `docs/` 配下で必要な画像・リンク・相対パス差分を確認する
- [x] `README.md` を刷新後の日本語版構成に合わせて更新する
- [x] `docs/README.ja.md` を刷新後の日本語版構成に合わせて更新する
- [x] `docs/README.en.md` を刷新後の日本語版構成に合わせて更新する
- [x] 差分確認で 3 ファイルの整合性を確認し、レビュー結果を記録する

### Review

- 3 ファイルとも先頭をロゴ + バッジ構成へ統一し、刷新後 `README.ja.md` に合わせて導入文・章立て・MCP 節を更新した
- `docs/README.ja.md` / `docs/README.en.md` は `assets/...` や `architecture/...` など `docs/` 配下向けの相対パスへ調整した
- `Documentation build` / `MCP追加メモ（UX Hardening）` / 旧 `Migration note` など、基準 README にない節は除去した
- `git diff` と `rg` で対象 3 ファイルに対する主要見出し・MCP 詳細・画像パス・不要節の有無を確認した

## Planning

- [x] issue 56 の本文とコメントを確認し、must-have / nice-to-have / 非ゴールを整理する
- [x] 既存の pipeline / shape / chart / render / MCP 実装を確認し、変更境界を特定する
- [x] `tasks/feature_spec.md` に `libreoffice` mode の仕様、型、fallback 方針を定義する
- [x] 実装順と検証方針をこの `tasks/todo.md` に落とし込む

## Phase 1: Public Contract

- [x] `ExtractionMode` を `light/libreoffice/standard/verbose` に拡張する
- [x] Python API (`extract`, `process_excel`, `StructOptions`, `extract_workbook`) の型と docstring を更新する
- [x] CLI `--mode` choices と help を更新する
- [x] MCP `ExtractRequest`, `server.py`, `docs/mcp.md` の mode 説明を更新する
- [x] `.xls` + `mode="libreoffice"` を早期バリデーションで拒否する

## Phase 2: Pipeline / Runtime

- [x] `resolve_extraction_inputs` に `libreoffice` の既定 include_* を追加する
- [x] `PipelineState` / `FallbackReason` に LibreOffice 用 reason を追加する
- [x] pipeline の rich backend 選択を `light/com/libreoffice` 前提で整理する
- [x] LibreOffice session helper を追加し、headless 起動・一時 profile・timeout・cleanup を実装する
- [x] LibreOffice 不在時の fallback を cells/tables/print_areas/merged_cells 維持で実装する

## Phase 3: Shape / Connector

- [x] OOXML drawing helper を追加し、shape / connector / chart anchor 情報を読めるようにする
- [x] LibreOffice UNO から draw-page shapes を取得する backend を追加する
  - `LibreOfficeSession.extract_draw_page_shapes(...)` と bridge payload を実装する
  - `LibreOfficeRichBackend.extract_shapes(...)` が UNO draw-page 順を canonical source として使う
- [x] LibreOffice bridge payload に draw-page shape と connector direct-ref を追加する
- [x] LibreOffice backend で UNO draw-page payload を shape metadata と connector 解決に統合する
- [x] non-connector shape のみシート内連番 `id` を振る仕様を実装する
- [x] connector 解決を `OOXML explicit ref -> UNO direct ref -> geometry heuristic` の優先順で実装する
  - OOXML connector match が取れない場合でも UNO `StartShape/EndShape` を使って begin/end を復元する
  - direct ref 不可時のみ geometry heuristic に落とす回帰 test を追加する
- [x] `BaseShape` metadata (`provenance`, `approximation_level`, `confidence`) を追加し、COM / LibreOffice 両経路で埋める

## Phase 4: Chart

- [x] OOXML / openpyxl から chart の semantic 情報を抽出する helper を追加する
- [x] LibreOffice UNO から chart geometry 候補を取得する
- [x] OOXML chart と UNO geometry を順序で pairing し、geometry を `Chart` に反映する
- [x] UNO geometry が無い場合は openpyxl anchor を geometry fallback として使う
- [x] `Chart` metadata (`provenance`, `approximation_level`, `confidence`) を追加する

## Phase 5: Verification

- [x] mode validation の unit test を追加する
- [x] `.xls` reject の unit test を追加する
- [x] `sample/flowchart/sample-shape-connector.xlsx` を使った connector graph 回帰 test を追加する
- [x] `sample/basic/sample.xlsx` を使った chart extraction 回帰 test を追加する
- [x] LibreOffice unavailable fallback の unit test を追加する
- [x] 必要なら `pytest.mark.libreoffice` の optional smoke test を追加する
- [x] `uv run pytest` または対象 test を実行して結果を確認する
- [x] `uv run task precommit-run` を実行し、ruff / mypy / format 系の問題が無いことを確認する

## Phase 6: Documentation

- [x] README.md / README.ja.md の mode 説明を更新する
- [x] contributor / architecture / release notes を更新する
- [x] `libreoffice` mode が best-effort であり strict subset ではないことを明記する
- [x] rendering と auto page-break が v1 対象外であることを明記する

## Phase 7: CLI LibreOffice Follow-up

- [x] `tasks/feature_spec.md` に CLI / process API の互換性ルールを追記する
- [x] `engine` / `integrate` に `libreoffice` 非対応オプションの共通バリデーションを追加する
- [x] CLI help を更新し、`--pdf` / `--image` / `--auto-page-breaks-dir` の制約を明記する
- [x] `process_excel(...)` / `ExStructEngine.process(...)` / `extract_workbook(...)` の docstring と例を新仕様に揃える
- [x] CLI / API / extract レイヤーの異常系 test を追加する
- [x] README / CLI Guide / API docs / test requirements を今回の契約に揃える

## Review

- 2026-03-06 draw-page / connector follow-up:
  - LibreOffice bridge に `--kind draw-page` を追加し、`DrawPage` 由来の shape / connector payload を取得可能にした
  - `extract_shapes(mode="libreoffice")` は UNO draw-page 順を canonical source にしつつ、OOXML を type / arrowhead / explicit ref 補完に限定して使う
  - connector 解決順を `OOXML explicit ref -> UNO direct ref -> geometry heuristic` に固定し、UNO-only と explicit-priority の unit test を追加した
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py tests/core/test_mode_output.py -k libreoffice -q` と `RUN_LIBREOFFICE_SMOKE=1` 付き smoke、`uv run task precommit-run` を通した
- 2026-03-06 chart geometry follow-up:
  - LibreOffice 同梱 Python bridge subprocess を追加し、`sheet.getCharts()` + `DrawPage` `OLE2Shape` から chart geometry 候補を取得
  - OOXML chart name / `PersistName` 一致を優先し、残差のみ順序 pairing する `libreoffice` mode の chart geometry 反映を実装
  - `RUN_LIBREOFFICE_SMOKE=1` 付き smoke test で chart geometry が 0 埋めではなく UNO geometry になることを確認

- 2026-03-06 follow-up:
  - `pytest.mark.libreoffice` smoke test と `RUN_LIBREOFFICE_SMOKE=1` gate を追加
  - LibreOffice runtime あり環境で smoke test を実行して通過
- 2026-03-06 CLI libreoffice follow-up:
  - `src/exstruct/constraints.py` を追加し、`mode="libreoffice"` と PDF/PNG rendering / auto page-break export の非対応組み合わせを `ConfigError` で早期拒否する共通ガードを実装
  - `extract_workbook(...)` / `ExStructEngine.process(...)` / CLI help / README / `docs/cli.md` / `docs/api.md` / `docs/agents/TEST_REQUIREMENTS.md` を同じ契約に揃えた
  - `uv run pytest tests/core/test_mode_output.py tests/cli/test_cli.py tests/backends/test_auto_page_breaks.py -q` は `28 passed, 1 skipped` で通過したが、既存 COM テスト由来の Windows COM 例外ログが終了後 stderr に出る現象は残っている
  - `uv run pytest tests/core/test_mode_output.py -k "libreoffice or process_excel_rejects or extract_workbook_rejects" tests/cli/test_cli.py -k "libreoffice" tests/backends/test_auto_page_breaks.py -k "libreoffice or extract_rejects_auto_page_break_flag" -q` は `8 passed`
  - `uv run task precommit-run` は ruff / ruff-format / mypy すべて通過

- 実装状態: 完了
- この時点で完了済み:
  - issue 56 の仕様整理
  - `feature_spec.md` 作成
  - 実装タスク分解
  - public contract / pipeline fallback / OOXML helper / best-effort backend 実装
  - mode / fallback / sample regression / metadata / docs 更新
  - `uv run pytest` 対象群と `uv run task precommit-run` の通過確認
- 実装完了条件:
  - public API / CLI / MCP の mode 追加が揃っている
  - connector graph と chart の best-effort 抽出が動く
  - 既存 COM 挙動に回帰がない
  - test / precommit-run が通る
- 主なリスク:
  - UNO API の環境差
  - connector heuristic の誤接続
  - chart geometry の pairing ずれ

## 2026-03-07 LibreOffice bridge compatibility probe follow-up

### Planning

- [x] `tasks/feature_spec.md` の互換性契約に合わせて、LibreOffice bridge 用 Python 候補の受け入れ条件を `bridge 実行可能` ベースへ更新する
- [x] `src/exstruct/core/_libreoffice_bridge.py` に no-op probe 用の `--probe` を追加し、UNO socket や workbook access を伴わない実行経路を用意する
- [x] `src/exstruct/core/libreoffice.py` の `_python_supports_libreoffice_bridge(...)` を、UNO import だけでなく bundled bridge probe 実行まで検証する実装へ変更する
- [x] `_resolve_python_path(...)` が system Python 候補を選ぶ際、`uno` import は通るが bridge 実行は失敗する候補を reject するようにする
- [x] `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` 指定時も同じ probe で fail-fast し、遅延 `SyntaxError` ではなく明確な incompatible runtime error を返す
- [x] `tests/core/test_libreoffice_backend.py` に、`uno` import success / bridge `SyntaxError` failure の false positive を防ぐ regression test を追加する
- [x] `tests/conftest.py` の LibreOffice runtime smoke gate も、必要なら bridge 実行可能性を反映した判定に揃える
- [x] `README.md` / `README.ja.md` / `docs/agents/TEST_REQUIREMENTS.md` の「compatible system Python」説明を probe ベースの表現へ揃える
- [x] `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py -q` と `uv run task precommit-run` で検証する

### Review

- 受け入れ条件:
  - system Python 自動検出は、bridge probe 成功済み候補だけを採用する
  - Debian 11 / Ubuntu 20.04 / WSL 想定の Python 3.8 / 3.9 + `python3-uno` false positive を regression test で再現し、以後は reject される
  - 明示 override も extraction 実行前に互換性エラーとして失敗し、`_run_bridge(...)` の `SyntaxError` まで遅延しない
  - smoke / docs / task 記述が「UNO import 可能」ではなく「bundled bridge 実行可能」で整合する
- 実装:
  - `_libreoffice_bridge.py` に internal `--probe` を追加し、`PropertyValue` と bridge 定数だけを解決する no-op 経路を用意した
  - `_python_supports_libreoffice_bridge(...)` は bundled bridge の `--probe` 実行で判定し、bundled / system Python 候補の両方で同じ互換性条件を使うようにした
  - `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` は同じ probe を必須にし、失敗時は `LibreOfficeUnavailableError` として fail-fast するようにした
  - `tests/conftest.py` の LibreOffice smoke gate も probe 失敗を unavailable 扱いに寄せた
  - README / 日本語 README / test requirements を probe ベースの説明へ更新した
- 追加検証:
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_pipeline_fallbacks.py tests/test_conftest_libreoffice_runtime.py -q` -> `22 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
## 2026-03-08 Linux LibreOffice CI smoke gate

### Planning

- [x] `tasks/feature_spec.md` に Linux CI smoke contract を追加し、`mode="libreoffice"` の required runtime gate を仕様化する
- [x] `.github/workflows/pytest.yml` に `libreoffice-linux-smoke` job を追加し、既存 unit matrix から分離する
- [x] `ubuntu-24.04` 上で `libreoffice` / `python3-uno` を導入して `RUN_LIBREOFFICE_SMOKE=1` で smoke test を実行する
- [x] smoke job は `tests/core/test_libreoffice_smoke.py` を対象にし、coverage upload とは分離する
- [x] README / README.ja / `docs/agents/TEST_REQUIREMENTS.md` に Linux required smoke job の運用を追記する
- [x] 変更内容の最終確認と検証コマンド結果を Review に反映する

### Review

- GitHub Actions に required Linux smoke job `libreoffice-linux-smoke` を追加し、`ubuntu-24.04` で LibreOffice runtime を導入して `pytest tests/core/test_libreoffice_smoke.py -m libreoffice` を実行する構成にした
- 既存の Linux/Windows unit matrix と Codecov upload はそのまま維持し、LibreOffice smoke は別 job に分離した
- `tasks/feature_spec.md` に Linux CI smoke contract を追加し、fallback 成功ではなく rich extraction 成功を merge 条件にする方針を明文化した
- README / 日本語 README / test requirements に required Linux smoke job と実行コマンドを追記した
- 追加検証:
  - `python -c "import yaml; yaml.safe_load(open('.github/workflows/pytest.yml', encoding='utf-8'))"` 相当の YAML parse -> `yaml-ok`
  - `uv run pytest tests/test_conftest_libreoffice_runtime.py -q` -> `1 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`

## 2026-03-09 PR #76 additional review triage

### Planning

- [x] `gh` で PR #76 の最新 review thread を再取得し、未 resolve 指摘を一覧化する
- [x] 各指摘を `feat/libreoffice-mode` の現行コードと `main...HEAD` 差分に照らして妥当性判定する
- [x] 採用する follow-up / 採用しない指摘を `tasks/feature_spec.md` に反映する
- [x] `tests/core/test_libreoffice_smoke.py` の confidence assertion を smoke 向けの高レベル検証へ緩める
- [x] `src/exstruct/core/backends/libreoffice_backend.py` の probe-only `_ensure_runtime()` を整理し、partial-success contract を維持したまま redundant session startup を削減する
- [x] `tests/core/test_mode_output.py` / `tests/cli/test_cli.py` の不自然な docstring を修正する
- [x] `AGENTS.md` の scope 外削除をこの PR から外す
- [x] 対象 pytest と `uv run task precommit-run` を実行して follow-up を検証する

### Review

- `gh api graphql` で PR #76 の未 resolve review thread は 7 件を確認した
- 対応採用とした指摘
  - `tests/core/test_libreoffice_smoke.py:39`
    - `chart.confidence == 0.8` は smoke としては backend 実装定数への結合が強い
    - exact `0.8` は `tests/core/test_libreoffice_backend.py` の unit test に残し、smoke は range/assertion へ寄せる
  - `src/exstruct/core/backends/libreoffice_backend.py:149-185`
    - `_ensure_runtime()` は probe 専用 session を 1 回余分に起動しており、通常の shapes+charts 経路で合計 3 回 startup になる
    - ただし「`_runtime_checked=True` で二回目 failure が隠れる」というレビュー本文の後半は不正確で、実読込 failure は別 session で surfacing される
    - follow-up では redundant probe だけを除去し、shape success 後の chart failure を保持する契約は維持する
  - `tests/core/test_mode_output.py:344,360`
    - `process excel sheets dir output` は不自然
    - `c l i` は typo
  - `tests/cli/test_cli.py:344`
    - `c l i` は typo
  - `AGENTS.md:49`
    - LibreOffice rollout と無関係な section 2/3/4 の大規模削除であり、PR scope 指摘は妥当
- 今回は実装対象に含めない指摘
  - `.github/workflows/pytest.yml:75`
    - `defusedxml` 未導入という指摘は誤り。`pyproject.toml` の core dependency なので `pip install -e .[...]` で導入される
    - `pip` / `uv` の統一と `pytest-cov` cleanup は maintenance 論点であり、この PR の blocking issue とは見なさない
  - `mkdocs.yml:50`
    - README nav 削除と `docs/README.*` 削除は docs build broken ではない
    - 論点は docs 導線再編の説明不足 / PR scope であり、必要なら review comment で意図を補足して resolve する
- 実装結果
  - smoke test は `chart.confidence is not None` と `0.0 <= confidence <= 1.0` に変更し、exact `0.8` 契約は backend unit test に残した
  - `LibreOfficeRichBackend` から probe-only `_ensure_runtime()` を外し、draw-page と chart read の実セッションだけを起動するようにした
  - `tests/core/test_mode_output.py` / `tests/cli/test_cli.py` の docstring を読みやすい英文に修正し、`tests/core/test_pipeline.py` の残っていた自動生成風 docstring も同じ sweep で整理した
  - `AGENTS.md` は `origin/main...HEAD` 差分で落ちていた section 2/3/4 を復元した
- 検証
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_libreoffice_bridge.py tests/core/test_libreoffice_smoke.py tests/core/test_pipeline.py tests/core/test_mode_output.py tests/cli/test_cli.py -q` -> `122 passed, 2 skipped`
  - pytest 終了後に既知の Windows COM fatal exception ログは出るが、終了コードは 0 のまま
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`

## 2026-03-09 PR #76 Codacy command-injection triage

### Planning

- [x] `scripts/codacy_issues.py --pr 76 --min-level Warning` で PR #76 の Codacy 残件を再取得し、issue が 1 件だけ残っていることを確認する
- [x] `src/exstruct/core/libreoffice.py:825` の sink が `_run_bridge_probe_subprocess(...)` であることを確認し、現行 helper/test の契約を点検する
- [x] Codacy の `dangerous-subprocess-use-tainted-env-args` を false positive 寄りと判断しつつ、まずは suppression ではなく trust-boundary 明確化で解く方針を決める
- [x] probe helper の `subprocess.run(...)` から explicit `env=` を外し、UTF-8 強制を固定 argv オプションへ移す
- [x] probe helper 専用の regression test を追加し、fixed argv と `env` 非指定を確認する
- [x] 既存の probe env test を新契約へ更新し、互換性判定まわりの高レベル test が壊れていないことを確認する
- [x] `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_libreoffice_bridge.py -q` を実行する
- [x] `uv run task precommit-run` を実行する
- [x] push 後に `python scripts/codacy_issues.py --pr 76 --min-level Warning` を再実行し、同 issue が消えたことを確認する
- [ ] それでも Codacy が同じ sink を報告する場合だけ、対象 call site に rule-specific suppression を追加する

### Review

- 2026-03-09 時点の Codacy 出力は次の 1 件:
  - `Error | src/exstruct/core/libreoffice.py:825 | Semgrep_python.lang.security.audit.dangerous-subprocess-use-tainted-env-args.dangerous-subprocess-use-tainted-env-args | Security | Detected subprocess function 'run' with user controlled data.`
- line 825 は `_run_bridge_probe_subprocess(...)` の `subprocess.run(...)` に対応している。
- 現行コードの安全性評価:
  - `shell=False`
  - command string 組み立てなし
  - executable は `_validated_runtime_path(...)` を通す
  - bridge script は repository 内の固定 path
  - probe helper は workbook path を受け取らない
- 一方で、Codacy は次を taint と見ている可能性が高い:
  - `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` / `sys.executable` / `shutil.which(...)` 由来の `python_path`
  - `_build_subprocess_env(...)` が allowlist 経由で取り込む `PATH` などの inherited env
- 採用方針:
  - まず probe helper から explicit env 注入を外し、UTF-8 制御を argv 化して analyzer の taint 経路を減らす
  - それでも残る場合だけ narrow suppression を入れる
  - いきなり suppression だけで終わらせない
- 実装結果:
  - `_run_bridge_probe_subprocess(...)` は `env=` を渡さず、固定 argv `python -X utf8 _libreoffice_bridge.py --probe` を使うように変更した
  - `test_python_supports_libreoffice_bridge_uses_probe_command` を新 argv 形状へ更新した
  - `test_run_bridge_probe_subprocess_uses_fixed_utf8_args_without_env` を追加し、`env` 非指定と UTF-8 runtime option を固定した
- 検証結果:
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_libreoffice_bridge.py -q` -> `49 passed`
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
- 残タスク:
  - なし。push 後の `python scripts/codacy_issues.py --pr 76 --min-level Warning` は `total: 0` を返した

## 2026-03-09 PR #76 latest review + Codacy re-triage

### Planning

- [x] `gh api graphql` で PR #76 の最新 unresolved review thread を再取得する
- [x] duplicate review thread 3 件に返信し、元 thread へ集約する形で resolve する
- [x] `python scripts/codacy_issues.py --pr 76 --min-level Warning` を再実行し、残件 rule が `dangerous-subprocess-use-audit` に切り替わったことを確認する
- [x] 新規/残存 review 指摘を採用・非採用に再分類する
- [x] `src/exstruct/core/backends/libreoffice_backend.py::_resolve_direction()` に rotation-aware delta を適用する
- [x] `src/exstruct/core/ooxml_drawing.py::_extract_chart_series()` を全 chart node 走査へ拡張する
- [x] `src/exstruct/core/ooxml_drawing.py::_parse_connector_node()` の `headEnd/tailEnd` mapping を begin/end semantics に合わせて修正する
- [x] `src/exstruct/core/_libreoffice_bridge.py::_resolve_context()` を at-least-once attempt になる loop に直す
- [x] `tests/core/test_libreoffice_smoke.py` の `confidence == 0.8` を smoke 向け assertion に緩める
- [x] `src/exstruct/core/backends/libreoffice_backend.py::_ensure_runtime()` を整理して redundant startup を除去する
- [x] `tests/core/test_pipeline.py` / `tests/core/test_mode_output.py` / `tests/cli/test_cli.py` の不自然な docstring を修正する
- [x] `AGENTS.md` の PR scope 外削除を戻す
- [x] trusted subprocess helper 群に narrow `nosemgrep` suppression を追加する
- [x] 対象 pytest を実行する
- [x] `uv run task precommit-run` を実行する
- [x] push 後に `python scripts/codacy_issues.py --pr 76 --min-level Warning` を再実行する

### Review

- unresolved thread 再取得後、duplicate として閉じた thread:
  - `discussion_r2904696477` -> `discussion_r2901508431` に集約
  - `discussion_r2904696479` -> `discussion_r2901508430` に集約
  - `discussion_r2901522451` -> `discussion_r2901509039` に集約
- 現在の open review 論点:
  - `src/exstruct/core/backends/libreoffice_backend.py`
    - connector `direction` に `rotation` を反映していない
    - `_ensure_runtime()` が probe-only startup を 1 回余分に行う
  - `src/exstruct/core/ooxml_drawing.py`
    - combo chart の secondary chart node series を落としている
    - connector `headEnd/tailEnd` mapping が COM semantics と逆
  - `src/exstruct/core/_libreoffice_bridge.py`
    - `_resolve_context()` が no-attempt timeout になり得る
  - tests
    - smoke confidence 固定値
    - `test_pipeline.py` / `test_mode_output.py` / `test_cli.py` の不自然な docstring
  - repo scope
    - `AGENTS.md` の scope 外削除
- Codacy 最新結果:
  - `Error | src/exstruct/core/libreoffice.py:806 | Semgrep_python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit | Security | Detected subprocess function 'run' without a static string.`
- Codacy 方針:
  - 前回の trust-boundary 縮小で `tainted-env-args` は消えた
  - 残っているのは generic audit rule なので、trusted helper に限定した `nosemgrep` で対処する
  - `_spawn_trusted_subprocess(...)` と同じ rule id を、`_run_soffice_version_subprocess(...)` と bridge `subprocess.run(...)` helper 群にも揃える
- 実装結果:
  - `_resolve_direction()` は `_rotate_connector_delta(...)` を通した回転後ベクトルで方位を決めるようにし、heuristic endpoint 推定と `direction` の幾何学を一致させた
  - `_extract_chart_series()` は `plotArea` 配下の全 chart node を document order で走査し、combo chart の secondary series も保持するようにした
  - `_parse_connector_node()` は `headEnd -> begin_arrow_style`、`tailEnd -> end_arrow_style` に修正した
  - `_resolve_context()` は timeout が極端に短くても 1 回は `resolver.resolve(...)` を試す loop に直した
  - `_run_soffice_version_subprocess()` と bridge `subprocess.run(...)` helper 群に rule-specific `nosemgrep` を追加した
  - 回帰 test を `tests/core/test_libreoffice_backend.py` / `tests/core/test_libreoffice_bridge.py` に追加・更新した
- 検証:
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_libreoffice_bridge.py tests/core/test_libreoffice_smoke.py tests/core/test_pipeline.py tests/core/test_mode_output.py tests/cli/test_cli.py -q` -> `122 passed, 2 skipped`
  - pytest 終了後に既知の Windows COM fatal exception ログは出るが、pytest 自体の終了コードは 0
  - `uv run task precommit-run` -> `ruff / ruff-format / mypy passed`
  - push 後に 2 分待ってから `python scripts/codacy_issues.py --pr 76 --min-level Warning` を 1 回だけ再実行し、`total: 0` を確認した
  - `gh api graphql` 再確認で PR #76 の review thread は全件 `isResolved: true` になった

## 2026-03-10 Windows LibreOffice CI smoke gate

### Planning

- [x] 既存の Linux LibreOffice smoke job / README / test requirements を確認し、Windows smoke 追加の最小変更方針を決める
- [x] `.github/workflows/pytest.yml` に Windows 専用 job `libreoffice-windows-smoke` を追加する
- [x] Windows hosted runner で `choco install libreoffice-fresh`、`RUN_LIBREOFFICE_SMOKE=1`、`FORCE_LIBREOFFICE_SMOKE=1`、`EXSTRUCT_LIBREOFFICE_PATH` を設定する
- [x] install 後に `soffice.exe` の存在確認と `--version` 実行を入れて fail-fast 化する
- [x] README / README.ja / `docs/agents/TEST_REQUIREMENTS.md` / `tasks/feature_spec.md` に Windows smoke job の契約を追記する
- [x] YAML parse と既存 pytest / pre-commit による最終確認を行う

### Review

- `.github/workflows/pytest.yml` に `windows-2025` 固定の `libreoffice-windows-smoke` job を追加した
- Windows job は Linux smoke と同じく unit matrix から独立させ、LibreOffice smoke だけを担当する構成にした
- runtime 導入は issue 提案どおり `choco install libreoffice-fresh -y --no-progress` を使い、`EXSTRUCT_LIBREOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe` を明示する
- install 後に `Test-Path` と `soffice.exe --version` を実行し、runtime 不在/破損を smoke 実行前に fail-fast させる
- README / 日本語 README / test requirements / feature spec を Windows smoke job 反映に更新した
- 検証:
  - `python -c "import yaml; yaml.safe_load(open('.github/workflows/pytest.yml', encoding='utf-8')); print('yaml-ok')"` -> `yaml-ok`
  - `python -m pytest tests/test_conftest_libreoffice_runtime.py -q` -> `3 passed`
  - `python -m pre_commit run -a` -> `ruff / ruff-format / mypy passed`

## 2026-03-10 Windows LibreOffice CI failure follow-up

### Planning

- [x] GitHub Actions の失敗 run / job / logs を確認し、失敗点を `libreoffice-windows-smoke` に絞る
- [x] `tests/conftest.py` と `src/exstruct/core/libreoffice.py` を確認し、Windows runtime unavailable の原因を切り分ける
- [x] bundled Python auto-detection が `python-core-*` 配下を探索していないギャップを埋める
- [x] `tests/core/test_libreoffice_backend.py` に Windows layout の regression test を追加する
- [x] 対象 pytest と `python -m pre_commit run -a` を実行する

### Review

- GitHub Actions run `22904348826` の job `libreoffice-windows-smoke` だけが failure で、`Install LibreOffice runtime` と `Verify LibreOffice runtime` は success、失敗は `tests/conftest.py::_has_libreoffice_runtime()` による setup error だった
- 既存の `_resolve_python_path(...)` は `program/python.exe` 等の直下候補しか見ておらず、Windows LibreOffice install の `python-core-*` 配下 bundled Python を見逃していた
- `src/exstruct/core/libreoffice.py` に `_bundled_python_candidates(program_dir)` を追加し、従来の直下候補に加えて `python-core-*/python.exe` / `python-core-*/bin/python.exe` を探索するようにした
- `tests/core/test_libreoffice_backend.py` に Windows `python-core-*` layout の回帰 test を追加した
- 検証:
  - `python3 -m pytest tests/core/test_libreoffice_backend.py -q` -> `48 passed`
  - `python3 -m pytest tests/test_conftest_libreoffice_runtime.py -q` -> `3 passed`
  - `python3 -m pre_commit run -a` -> `ruff / ruff-format / mypy passed`

## 2026-03-10 Windows LibreOffice CI failure second follow-up

### Planning

- [x] 最新の failing workflow run / job / logs を確認し、まだ残っている failure point を特定する
- [x] runtime gate と bridge probe の実装差分を確認し、Windows hosted runner で不足している条件を見つける
- [x] probe subprocess に必要な最小 env だけを forward するよう修正する
- [x] 関連 regression test を更新し、targeted pytest と pre-commit を再実行する

### Review

- 最新 run `22905085870` でも failure は `libreoffice-windows-smoke` の setup error で、`soffice.exe --version` 成功後に `tests/conftest.py::_has_libreoffice_runtime()` が `False` になっていた
- `_run_bridge_probe_subprocess(...)` だけが `_build_subprocess_env(...)` を使っておらず、Windows hosted runner の LibreOffice Python / UNO probe に必要な runtime env を欠く可能性があった
- `src/exstruct/core/libreoffice.py` で probe subprocess にも allowlisted env + `PYTHONIOENCODING=utf-8` を渡すようにし、bridge handshake / extraction と整合させた
- `tests/core/test_libreoffice_backend.py` の probe subprocess test を、env を forward しつつ allowlist 外 env を漏らさない契約に更新した
- 検証:
  - `python3 -m pytest tests/core/test_libreoffice_backend.py -q` -> `48 passed`
  - `python3 -m pytest tests/test_conftest_libreoffice_runtime.py -q` -> `3 passed`
  - `python3 -m pre_commit run -a` -> `ruff / ruff-format / mypy passed`

## 2026-03-10 Windows LibreOffice CI failure third follow-up

### Planning

- [x] 最新の workflow run / logs を再確認し、failure が依然として runtime gate setup にあることを確認する
- [x] bundled Python auto-detection と Windows workflow assumptions を見直し、CI 固有の残ギャップを特定する
- [x] workflow で bundled Python path を明示 discovery して `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` を渡す最小修正を入れる
- [x] YAML parse と関連 validation を実行し、変更を push する

### Review

- 最新 run `22905850626` でも failure は `libreoffice-windows-smoke` の setup error で、`soffice.exe --version` 成功後に `tests/conftest.py::_has_libreoffice_runtime()` が `False` になっていた
- code 側の bundled Python auto-detection は `python.exe` / `python.bin` / `python` / `python-core-*` 系を探索するが、workflow では `EXSTRUCT_LIBREOFFICE_PATH` しか固定しておらず、hosted runner 上の actual bundled Python path は runtime gate に委ねられていた
- `.github/workflows/pytest.yml` に `Discover LibreOffice bundled Python` step を追加し、LibreOffice install 後に bundled Python executable を探索して `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` として後続 step に渡すようにした
- bundled Python が見つからない場合は、smoke 実行前に program directory listing を出して fail-fast するようにし、次回以降の CI 調査を容易にした
- `Verify LibreOffice runtime` step でも discovered Python path の存在確認を追加した
- 検証:
  - `python3 - <<'PY' ... yaml.safe_load('.github/workflows/pytest.yml') ...` -> `yaml-ok`
  - `python3 -m pytest tests/test_conftest_libreoffice_runtime.py -q` -> `3 passed`
  - `python3 -m pre_commit run -a` -> `ruff / ruff-format / mypy passed`

## 2026-03-10 Windows LibreOffice CI failure fourth follow-up

### Planning

- [x] 最新の `libreoffice-windows-smoke` log を再確認し、workflow discovery 後も runtime gate で落ちていることを確認する
- [x] discovered `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` が pytest step に届いていることを確認し、残差分が bridge probe subprocess 条件にあると切り分ける
- [x] bridge subprocess を bundled Python parent directory を `cwd` にして起動する最小コード修正を入れる
- [x] focused tests と validation を実行し、結果を記録する

### Review

- 最新 run `22906728072` では `Discover LibreOffice bundled Python` が成功し、`EXSTRUCT_LIBREOFFICE_PYTHON_PATH=C:\\Program Files\\LibreOffice\\program\\python.exe` が pytest step に渡っていた
- それでも `tests/conftest.py::_has_libreoffice_runtime()` が `False` になっていたため、残障害は path discovery ではなく bridge probe subprocess 側だと切り分けた
- `src/exstruct/core/_libreoffice_bridge.py` は module import 時点で `uno` を import するため、Windows bundled Python は probe / handshake / extraction すべてで LibreOffice program directory を working directory として持つ必要がある
- `src/exstruct/core/libreoffice.py` に `_bridge_subprocess_cwd(...)` を追加し、bridge subprocess 3系統 (`--probe`, `--handshake`, extraction) で `cwd=_validated_runtime_path(python_path).parent` を使うようにした
- `tests/core/test_libreoffice_backend.py` の focused test を更新し、probe / handshake / extraction が同じ `cwd` contract を使うことを検証した
- 検証:
  - `python3 -m pytest tests/core/test_libreoffice_backend.py -q` -> `48 passed`
  - `python3 -m pytest tests/test_conftest_libreoffice_runtime.py -q` -> `3 passed`
  - `python3 -m pre_commit run -a` -> `ruff / ruff-format / mypy passed`

## 2026-03-10 PR79 LibreOffice Windows smoke failure investigation

### Planning

- [x] 失敗ログと `tests/conftest.py` の runtime gate を照合し、Windows CI false-negative の経路を特定する
- [x] `tasks/feature_spec.md` で runtime gate timeout fallback 仕様を明文化する
- [x] `tests/conftest.py` の LibreOffice 判定を timeout 耐性付きに修正する
- [x] 回帰テストを追加して timeout fallback 契約を固定する
- [x] 対象 pytest を実行して挙動を検証する

### Review

- `tests/conftest.py::_has_libreoffice_runtime()` の `soffice --version` probe で `TimeoutExpired` が起きた場合、即 `False` ではなく `LibreOfficeSession.from_env()` の短命セッション probe を 1 回実施するように変更した。
- fallback session が `LibreOfficeUnavailableError` を返す場合のみ unavailable (`False`) とし、予期しない例外は従来どおり surfacing して fail-fast を維持した。
- `tests/test_conftest_libreoffice_runtime.py` に timeout fallback の成功/失敗ケース回帰テストを追加し、Windows CI で起きうる初期タイムアウトの false-negative を防ぐ契約を固定した。
- 検証:
  - `pytest tests/test_conftest_libreoffice_runtime.py -q` -> `5 passed`
  - `uv run ruff check tests/conftest.py tests/test_conftest_libreoffice_runtime.py` -> pass
  - `uv run mypy tests/conftest.py tests/test_conftest_libreoffice_runtime.py` -> pass
  - `uv run task precommit-run` は pre-commit hook の remote fetch (`astral-sh/ruff-pre-commit`) が `CONNECT tunnel failed, response 403` で失敗（環境制約）


## 2026-03-10 PR79 follow-up retry hardening

### Planning

- [x] CI再失敗ログを前回修正との差分観点で再分析する
- [x] slow probe retry 方針を `tasks/feature_spec.md` に追記する
- [x] `tests/conftest.py` に version probe retry を実装する
- [x] retry 契約の回帰テストを追加する
- [x] pytest + lint/type check を再実行する
- [x] lessons を更新する

### Review

- `tests/conftest.py` の LibreOffice runtime gate を再調整し、`soffice --version` の 5 秒 probe timeout 後に 30 秒で 1 回再試行するようにした。
- 再試行が成功した場合は `True` を返し、重い session fallback (`LibreOfficeSession.from_env()`) は呼ばない。
- 再試行も timeout の場合のみ既存の session fallback に委譲し、`LibreOfficeUnavailableError` は `False`、予期しない例外は surfacing を維持した。
- `tests/test_conftest_libreoffice_runtime.py` に、初回 timeout -> 再試行成功で `True` を返し session fallback を通らない回帰テストを追加した。
- `tasks/lessons.md` に、Windows cold-start probe は short-timeout single-shot にせず long-timeout retry 層を先に置く学びを追記した。
- 検証:
  - `pytest tests/test_conftest_libreoffice_runtime.py -q` -> `6 passed`
  - `uv run ruff check tests/conftest.py tests/test_conftest_libreoffice_runtime.py` -> pass
  - `uv run mypy tests/conftest.py tests/test_conftest_libreoffice_runtime.py` -> pass
  - `uv run task precommit-run` は pre-commit hook remote fetch が `CONNECT tunnel failed, response 403` で失敗（環境制約）


## 2026-03-11 PR #79 Windows LibreOffice smoke CI fix

### Planning

- [x] 現行 `libreoffice-windows-smoke` workflow と runtime 正規化実装を確認する
- [x] Windows で `soffice.com` 優先となるよう runtime path 正規化と workflow を修正する
- [x] 回帰テストを追加し、対象 pytest を実行して検証する
- [x] 変更内容を自己レビューし、commit/PR メッセージを作成する

### Review

- `src/exstruct/core/libreoffice.py` で runtime path 正規化時に Windows の `soffice.exe` を `soffice.com` 優先へ自動変換する helper を追加した。
- `src/exstruct/core/libreoffice.py` の `_which_soffice()` は `soffice.com` を探索候補に追加し、検出 path を正規化して返すようにした。
- `.github/workflows/pytest.yml` の Windows smoke job は `soffice.com` を既定にしつつ、discover step で `.com` 優先 / `.exe` fallback を明示し `EXSTRUCT_LIBREOFFICE_PATH` を再設定するよう変更した。
- 同 workflow の verify step で `--version` 実行後 `$LASTEXITCODE` を確認し、非ゼロを即失敗にするようにした。
- `tests/core/test_libreoffice_backend.py` に runtime path 正規化の Windows/非 Windows 回帰テストを追加した。
