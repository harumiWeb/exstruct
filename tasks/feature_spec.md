# Feature Spec

## 2026-03-10 PR #76 review + Codacy follow-up

### Issue

- PR #76 に 2026-03-10 時点で未 resolve review thread が 2 件残っている。
  - `src/exstruct/core/backends/libreoffice_backend.py:_match_by_name_then_order()`
    - name match 後の order fallback が `len(remaining_snapshot_indexes) == len(unused)` 条件で止まり、件数不一致時の部分一致を落としている。
  - `src/exstruct/core/backends/libreoffice_backend.py:_direction_from_box()`
    - OOXML 向き情報が無い場合に UNO bbox (`width/height`) から heading を推定しており、左向き/上向き系 connector で誤方向を返しうる。
- Codacy check も `src/exstruct/core/libreoffice.py:13` の `import subprocess` に対する `Bandit_B404` notice を 1 件返し、PR status を `action_required` にしている。

### Accepted follow-ups

- `_match_by_name_then_order()` の order fallback は件数一致を要求しない。
  - name match で消費した残差については `zip(..., strict=False)` の範囲で部分順序対応を許す。
  - これにより、OOXML-only / UNO-only の余剰要素が片側にあっても、相対順序で合わせられる候補は metadata を取り戻せる。
- connector `direction` は UNO bbox から推定しない。
  - 優先順位は `OOXML direction_dx/direction_dy`。
  - それが無い/ゼロ長の場合は、解決済み `begin_id/end_id` の shape geometry から direction を導く。
  - 解決済み endpoint geometry も無い場合は `None` を返し、誤った heading をでっち上げない。
- `src/exstruct/core/libreoffice.py` の `import subprocess` には `Bandit_B404` 向けの narrow suppression comment を付与する。
  - 対象は import 行のみ。
  - subprocess 利用自体は既存の validated local runtime/process management に限定される契約を維持する。

### Verification

- `tests/core/test_libreoffice_backend.py` に次の regression tests を追加/更新する。
  - `_match_by_name_then_order()` が件数不一致でも partial order fallback を適用する test
  - zero-length OOXML delta が resolved endpoint geometry に fallback する test
  - OOXML metadata も resolved endpoint geometry も無いとき direction が `None` になる test
- `uv run pytest tests/core/test_libreoffice_backend.py -q`
- `uv run task precommit-run`

## 2026-03-08 LibreOffice subprocess Codacy false-positive follow-up

### Issue

- `src/exstruct/core/libreoffice.py` の `subprocess.run(...)` helper に対して、Codacy が `Command Injection` (`dangerous-subprocess-use-audit`, `dangerous-subprocess-use-tainted-env-args`) を継続して報告している。
- 現行実装は `shell=False` + argv list + allowlisted env であり、実際のリスクは command string injection ではなく、静的解析が汎用 `args: Sequence[str]` helper 越しの trust boundary を追えていない点にある。

### Accepted follow-ups

- 汎用 `args: Sequence[str]` を受ける `_run_trusted_subprocess(...)` は廃止し、用途別の専用 helper へ分解する。
  - `soffice --version` probe
  - bridge `--probe`
  - bridge extraction (`--file`, `--kind`)
  - bridge handshake (`--handshake`)
- 各 helper は固定 argv 構造を内部で組み立て、可変入力は型付き引数として受ける。
  - executable path は `_validated_runtime_path(...)`
  - bridge script path は bundled local file
  - workbook path は `_subprocess_path_arg(...)` で単一 argv 要素として渡す
  - env は `_build_subprocess_env(...)` の allowlist を維持する
- `shell=False` と list argv は維持し、`shlex.escape()` のような shell 向け escaping は導入しない。
  - 理由: 本コードは shell command string を組み立てておらず、escaping を追加しても security posture は改善しないため
- 既存 public error contract は維持する。
  - bridge extraction timeout/error の例外文言
  - handshake failure の例外文言
  - Python bridge probe incompatibility detail
- regression tests:
  - bridge extraction helper が固定 argv / allowlisted env を使うこと
  - bridge probe helper が固定 argv / allowlisted env を使うこと
  - handshake helper が固定 argv / allowlisted env を使うこと
  - 既存の bridge extraction / invalid JSON / runtime-unavailable error mapping が変わらないこと

## 2026-03-08 PR #76 unresolved thread follow-up

### Issue

- PR #76 には 2026-03-08 時点で未 resolve thread が 2 件残っている。
  - `src/exstruct/core/backends/libreoffice_backend.py:539-541`
    - connector heuristic endpoint matching が OOXML の `rotation` を無視し、`left + dx`, `top + dy` の未回転ベクトルで begin/end を推定している。
  - `src/exstruct/core/pipeline.py:1009-1010`
    - LibreOffice rich extraction で `extract_shapes()` と `extract_charts()` を同じ `try` に入れているため、chart 側だけの failure でも shape artifact まで捨てて cells-only fallback に落ちる。

### Accepted follow-ups

- LibreOffice connector endpoint heuristic は OOXML connector の回転を反映してから begin/end 候補点を作る。
  - `direction_dx/direction_dy` は connector local axis のベクトルとして扱い、`rotation` がある場合は sheet 座標へ回転変換してから `left/top` に加算する。
  - `rotation` が未設定なら現行どおり未回転ベクトルを使う。
  - `(dx, dy) == (0, 0)` のときは既存 contract を維持し、OOXML endpoint heuristic には使わず UNO bounding box fallback を使う。
  - regression test:
    - direct endpoint refs が無く、`rotation=90` の connector でも heuristic begin/end が期待 shape に張り付くこと
    - `rotation=None` の既存 heuristic が変わらないこと
- LibreOffice pipeline fallback は rich artifact の部分成功を保持する。
  - `resolve_rich_backend(...)` failure は従来どおり `LIBREOFFICE_UNAVAILABLE` または `LIBREOFFICE_PIPELINE_FAILED` fallback に落としてよい。
  - ただし backend 解決後は `extract_shapes()` と `extract_charts()` を独立して扱い、片方が成功済みなら成功した artifact を保持したまま workbook を組み立てる。
  - chart extraction failure 時は extracted shapes を保持し、shapes extraction failure 時は fallback reason を維持しつつ charts だけを破棄してよい。
  - workbook build では `include_rich_artifacts=True` を使い、空 dict の rich artifact は単に空配列として扱う。
  - regression test:
    - shapes 成功 + charts failure で workbook に shapes が残り、`fallback_reason == LIBREOFFICE_PIPELINE_FAILED` になること
    - charts 成功 + shapes failure の場合は shapes/charts とも空で fallback になること

## 2026-03-08 PR #76 follow-up triage

### Issue

- PR #76 (`Add libreoffice extraction mode`) には、CI failure、Codacy 指摘、Codex/CodeRabbit review 指摘が複数残っている。
- 2026-03-08 時点の GitHub Actions failure は test failure ではなく、`test (ubuntu-latest, 3.12)` の coverage `78.62% < 80%` によるもの。
- Codacy は PR #76 に対して 1 件のみ `Error` を返しており、`src/exstruct/core/ooxml_drawing.py` の `xml.etree.ElementTree` 利用に対する XML hardening 指摘である。

### Goal

- 実害または公開契約の不整合がある指摘だけを follow-up scope に採用する。
- 採用した指摘は、仕様・検証観点・運用タスクをこの spec/todo に明文化する。
- スタイル提案や大規模リファクタ提案は、現時点で不具合根拠が無い限り scope から外す。

### Accepted follow-ups

- `ExStructEngine.process(...)` で per-call 指定した `auto_page_breaks_dir` / `include_auto_page_breaks` 相当の override は、validation と extraction の両方で同じ値を使う。
  - 現状は validation 時のみ per-call override を考慮し、`extract()` 呼び出しには伝播しないため、通常モードでも抽出対象と出力先が不整合になりうる。
  - follow-up では `extract()` 側に明示 override を渡せるようにするか、同等の一貫した内部経路へ整理する。
- LibreOffice guardrail の unit test は、少なくとも以下を直接カバーする。
  - `pdf=True`
  - `image=True`
  - `pdf/image + auto_page_breaks_dir`
  - `.xls + mode="libreoffice"`
  - non-`libreoffice` passthrough
  - `validate_libreoffice_extraction_request(...)` の direct coverage
- `docs/cli.md` と `docs/api.md` は shipped surface に一致させる。
  - `docs/cli.md` に `--include-backend-metadata` を追加する。
  - `docs/api.md` の重複した旧 `to_json(...)` signature を削除する。
- `libreoffice-linux-smoke` job は workflow 定義だけでなく merge gate としても一貫して扱う。
  - workflow step は skip を許容せず、skip が発生したら job failure にする。
  - repo ruleset / branch protection に required status check を追加するまでは、「required job」は repository contract として未成立であることを明記する。
- LibreOffice connector reconstruction は、片側だけ OOXML direct resolve できた場合でも残り片側を UNO / heuristic で継続解決する。
- connector direction 推定では、OOXML ベクトルが `(0, 0)` の場合を unknown 扱いとし、任意の方角に丸めない。
- OOXML drawing geometry は parent anchor (`oneCellAnchor` / `twoCellAnchor`) を canonical placement source とする。
  - child `xfrm` は width/height/rotation などの補助情報に限定する。
- OOXML chart series extraction は scatter/bubble 系の `c:xVal` / `c:yVal` も拾う。
- LibreOffice session 起動では、長寿命 `soffice` child の `stderr` pipe を未読のまま保持しない。
- `confidence` は public model / generated schema ともに `0.0 <= confidence <= 1.0` を強制する。
  - source of truth は model constraint とし、schema は生成物として同期する。
- OOXML parser hardening として、user-provided workbook 内 XML の parse は `defusedxml` ベースへ切り替える。

### Investigation follow-ups

- `snapshots` が存在するシートで OOXML-only shape / connector を append すべきかは、`UNO canonical order` 契約との整合を再確認してから決める。
- `_reserve_tcp_port()` の race は理論上成立するため、hold-open reservation へ変えるか、現行実装の許容理由を明文化するかを follow-up で判断する。

### 2026-03-08 follow-up decision

- `_reserve_tcp_port()` race は「port を完全予約する」方針ではなく、startup attempt 内の bounded retry で扱う。
  - 1 回の startup attempt (`isolated-profile` / `shared-profile`) は `port allocate -> soffice spawn -> socket wait` を最大 3 回まで再試行できる。
  - retry ごとに新しい ephemeral port を使い、短い backoff を入れる。
  - retry 対象は `socket startup timed out`、startup 中 exit、bind 失敗相当の stderr を含む port-collision 系の起動失敗。
  - 上限到達後は現在の aggregated startup failure に畳み込み、attempt 名と stderr detail を残す。
  - verification:
    - 1 回目失敗 / 2 回目成功の retry regression test
    - 上限到達時に failure detail を保持する regression test
- `draw-page snapshots` が存在する sheet では、OOXML-only shape / connector を v1 では append しない。
  - UNO draw-page snapshot order を canonical emitted order とする現行 contract を維持する。
  - OOXML は shape type / arrowhead / explicit refs / geometry の補助情報に限定し、snapshot 未対応要素を emitted list に追加しない。
  - 理由:
    - emitted `id` の安定性を崩さない
    - duplicate 判定を未定義のまま広げない
    - connector `begin_id/end_id` の解決 contract を保つ
    - `provenance` に OOXML-only を表す public contract を追加せずに済む
  - verification:
    - snapshot あり + unmatched OOXML shape/connector が emit されない unit test
  - observability:
    - append ではなく unmatched 件数を debug logging し、snapshot-backed emit contract を変えずに観測できるようにする

### 2026-03-08 post-push follow-up triage

#### Issue

- PR #76 の最新 GitHub Actions run `22814113410` は coverage ではなく test failure で落ちている。
  - `test (ubuntu-latest, 3.12)` の失敗は `tests/engine/test_engine.py::test_engine_process_normalizes_string_paths`
  - `windows-latest` matrix は fail-fast により cancel
- Codacy は PR #76 に対して `src/exstruct/core/libreoffice.py` の subprocess security 指摘を 4 件返している。
  - `_run_bridge(...)`
  - `_probe_soffice_runtime(...)`
  - `_probe_libreoffice_bridge_failure(...)`
  - `_start_soffice_startup_attempt(...)`
- 追加 review では、未 resolve thread が 2 件、非 thread の actionable comment が 2 件ある。
  - `_read_relationships()` が Relationship Type を捨てている
  - `tests/conftest.py` が broad `except Exception` で probe regression を skip に潰しうる
  - `docs/api.md` の CLI 例が `--include-backend-metadata` を落としている
  - `_merge_anchor_geometry(...)` が left/top を child transform 優先のまま

#### Accepted follow-ups

- `ExStructEngine.process(...)` は path 正規化後の実 extraction を `ExStructEngine.extract(...)` と同等の override 可能な seam から呼ぶ。
  - per-call validation (`mode`, `pdf`, `image`, `auto_page_breaks_dir`) は `process(...)` 側で先に確定してよい
  - ただし engine-level test が monkeypatch した `extract(...)` を bypass して実 pipeline に落ちる状態は contract regression とみなす
  - regression test:
    - `str` 入力 path
    - `pdf=True`
    - `image=True`
    - `export(...)` monkeypatch
    - 実ファイル内容に依存せず real pipeline へ入らないこと
- OOXML relationships は `target` だけでなく `type` を保持する structured model にする。
  - `_read_relationships(...)` は `dict[str, Relationship]` 相当を返す
  - call site は filename/prefix 推測ではなく relationship type URI で worksheet/drawing/chart/vml を判定する
  - 少なくとも `vmlDrawing*.vml` や non-worksheet relation を worksheet drawing と誤認しない
- `tests/conftest.py` の LibreOffice runtime probe は availability failure だけを skip 扱いにする。
  - `resolve_python_path(...)` では expected runtime-unavailable 系だけを `False` に落とす
  - `soffice --version` probe でも expected subprocess availability failure だけを `False` に落とす
  - unexpected exception は skip ではなく test failure として surfacing する
- `docs/api.md` の Python/CLI 対応例は `include_backend_metadata=True` に対して `--include-backend-metadata` を必ず対にする。
- OOXML drawing placement は parent anchor を sheet placement の source of truth とする。
  - `left/top` は anchor (`absoluteAnchor` / `oneCellAnchor` / `twoCellAnchor`) を優先する
  - child `xfrm` の `off` は anchor 欠落時の fallback に留める
  - `width/height/rotation/flip` は現行どおり child geometry を補助情報として使ってよい
- LibreOffice runtime subprocess は shell injection ではなく trust-boundary 明確化の issue として扱う。
  - すべて `shell=False` + argv list を維持する
  - 実行ファイル path は operator-configured runtime (`EXSTRUCT_LIBREOFFICE_PATH`, `EXSTRUCT_LIBREOFFICE_PYTHON_PATH`, bundled bridge path) として正規化・存在確認してから使う
  - workbook path を含む user data は argv element としてのみ渡し、command string 連結に使わない
  - static analyzer がなお誤検知する場合は、helper 集約後に最小スコープの suppression/comment で理由を明示する

#### Out of scope for this post-push follow-up

- `tests/core/test_libreoffice_backend.py` の `_pytest.monkeypatch.MonkeyPatch` を `pytest.MonkeyPatch` に寄せる style cleanup
- `_start_soffice_startup_attempt(...)` の helper 分割だけを目的にした complexity refactor

### 2026-03-08 coverage / Codacy / CodeRabbit follow-up triage

#### Issue

- PR #76 の最新 GitHub Actions run `22815422942` は test failure ではなく coverage gate failure で落ちている。
  - `test (ubuntu-latest, 3.12)` の `Run tests (non-COM suite)` は `783 passed, 3 skipped, 11 deselected`
  - failure reason は `Total coverage: 78.80% < 80%`
  - fail-fast により `ubuntu-latest, 3.11` / `windows-latest` matrix は `cancelled`
- ローカルで `uv run pytest -m "not com and not render" --cov=exstruct --cov-report=term-missing:skip-covered --cov-fail-under=0 -q` を再実行すると、coverage 低下は今回触ったモジュール群に偏っている。
  - `src/exstruct/core/libreoffice.py`: `67%`
  - `src/exstruct/core/backends/libreoffice_backend.py`: `79%`
  - `src/exstruct/core/ooxml_drawing.py`: `83%`
- Codacy は PR #76 に対して `src/exstruct/core/libreoffice.py` の subprocess security 指摘を 7 件返している。
  - `Bandit_B603` warning が `267`, `348`, `398`
  - Semgrep error が `267`, `268`, `348`, `398`
  - 現状コードは `shell=False` + argv list を維持しており、実体は command injection より trust-boundary の静的解析誤検知に近い
- CodeRabbit / review の新規 actionable 指摘は 2 系統ある。
  - `src/exstruct/engine.py`: `_process_extract_scope()` が `self.output.destinations.auto_page_breaks_dir` を直接 mutate し、immutable contract を破る
  - `src/exstruct/core/libreoffice.py`: `_reserve_tcp_port()` 後に別プロセスが同 port を取ると、`_wait_for_socket()` が「何かが listen している」だけで startup success と誤認しうる

#### Accepted follow-ups

- coverage gate は threshold 緩和ではなく targeted regression tests で回復する。
  - 追加テストは今回変更した low-coverage module に寄せる
  - 最優先は `libreoffice.py`, `libreoffice_backend.py`, `engine.py` の新規/変更分岐
  - verification は non-COM suite の coverage report を source of truth にする
- `ExStructEngine.process(...)` の per-call auto page-break override は engine-level seam を維持しつつ shared state mutation をなくす。
  - `process()` は引き続き engine-level `extract(...)` seam を通す
  - ただし override の伝播は `self.output.destinations` を一時 mutation するのではなく、explicit per-call parameter または private helper argument で表現する
  - `Instances are immutable` contract を壊さないことを優先する
  - regression test:
    - 同一 engine instance で異なる `auto_page_breaks_dir` を連続呼び出ししても leak しない
    - monkeypatch した `extract(...)` seam を bypass しない
    - `include_auto_page_breaks` の解決値が caller ごとに独立している
- LibreOffice startup success は「TCP accept できた」ではなく「期待する UNO bridge に実際に到達できた」で判定する。
  - `_wait_for_socket()` は引き続き初期 readiness probe として使ってよい
  - ただし success 確定前に、選択された Python runtime と bundled bridge で host/port に対する lightweight handshake を行う
  - handshake failure は port collision / wrong listener とみなし、startup retry に戻す
  - PID/port owner lookup のためだけに `psutil` / `lsof` 依存を追加しない
  - regression test:
    - foreign listener が accept するだけのケースを startup failure と判定する
    - handshake success 時だけ session enter が成立する
- Codacy subprocess findings は runtime design 自体の欠陥ではなく static-analysis 誤検知として扱い、最小スコープ suppression までを対応範囲に含める。
  - `shell=False` + argv list を維持する
  - executable path は `_validated_runtime_path(...)` を通した operator-configured runtime のみを使う
  - workbook / bridge path は単一 argv 要素として渡し、command string 連結は行わない
  - analyzer が helper 越しでも警告を維持する call site には、`B603` / Semgrep 向けの inline suppression/comment を最小範囲で付け、理由をコード上に残す
  - suppression は `src/exstruct/core/libreoffice.py` の該当 4 call site に限定する

#### Out of scope for this follow-up

- coverage threshold を `80` 未満へ下げる
- low-coverage module を `.coveragerc` / workflow で除外して見かけ上の coverage を上げる
- startup listener 検証のためだけに process ownership 依存 (`psutil`, `lsof`) を増やす
- subprocess 呼び出し全体の大規模 API redesign

### Out of scope for this follow-up

- `ShapeData` / `ChartData` を dataclass/Pydantic へ全面変更するリファクタ
- `LibreOfficeSession.load_workbook()` / `close_workbook()` の typed handle 化だけを目的にした API 再設計
- `normalize_path(...)` docstring のみを目的とした style fix
- `.xls + mode="libreoffice"` の例外型を `ValueError` から `ConfigError` へ変更すること
- `get_charts(..., mode=...)` の未使用引数に対する docstring-only 修正

## 2026-03-06 backend metadata output follow-up

- shape/chart backend metadata fields (`provenance`, `approximation_level`, `confidence`) remain part of the internal models.
- Serialized output treats these fields as opt-in to reduce token usage.
- A shared public flag `include_backend_metadata` controls emission for Python serialization helpers, `ExStructEngine` filters, CLI, and MCP extract requests.
- Default behavior:
  - CLI: `--include-backend-metadata` not specified -> metadata omitted
  - MCP: `options.include_backend_metadata=false`
  - Python serialization helpers: `include_backend_metadata=False`
- Backend extraction logic and schema field definitions remain unchanged; only output shaping changes.

## Issue

- Issue #56: LibreOffice backend / `libreoffice` mode 追加
- 目的: Excel COM に依存していた一部の抽出を、Linux / macOS / server / CI でも best-effort で実行できるようにする

## 背景

- 現状の rich extraction は `standard` / `verbose` で Excel COM に依存している。
- 非 COM 環境では `light` 相当のフォールバックしかなく、shape / connector / chart の構造情報が落ちる。
- issue 56 の要求は「COM 同等の厳密性」ではなく、「server-first の best-effort 抽出」を追加することにある。

## ゴール

- 新しい抽出モード `libreoffice` を追加する。
- `libreoffice` mode は `light` より多く、`standard` より少ない情報を返す。
- v1 では shape / connector / chart を対象にし、connector graph を最優先で復元する。
- 出力に provenance / confidence を持たせ、downstream が精度差を判定できるようにする。

## 非ゴール

- Excel COM と同一のレイアウト再現
- Excel の `DisplayFormat` 相当の見た目再現
- SmartArt の忠実な復元
- auto page-break の LibreOffice 版実装
- LibreOffice を使った PDF / PNG rendering の追加
- `.xls` の LibreOffice mode 対応
- `standard` / `verbose` での自動 LibreOffice フォールバック

## 公開仕様

### 1. モード

`ExtractionMode` は次の 4 値にする。

```python
ExtractionMode = Literal["light", "libreoffice", "standard", "verbose"]
```

各モードの意味は次のとおり。

| mode | 目的 | 主な出力 |
| --- | --- | --- |
| `light` | 最小・高速 | cells, table_candidates, print_areas |
| `libreoffice` | 非 COM 環境向け best-effort | `light` + merged_cells + shapes + connectors + charts |
| `standard` | 既定 | COM 利用可能なら既存の rich extraction |
| `verbose` | 最多情報 | `standard` + size / links / maps |

### 2. `libreoffice` mode の既定値

- `include_cell_links=False`
- `include_print_areas=True`
- `include_auto_page_breaks=False`
- `include_colors_map=False`
- `include_formulas_map=False`
- `include_merged_cells=True`
- `include_merged_values_in_rows=True`

### 3. `libreoffice` mode の対象拡張子

- 対象: `.xlsx`, `.xlsm`
- 非対象: `.xls`
- `.xls` を `mode="libreoffice"` で指定した場合は、処理開始前に `ValueError` を返す。
- エラーメッセージは「`.xls` is not supported in libreoffice mode; use COM-backed standard/verbose or convert to .xlsx`」相当の明確な文言にする。

### 4. `libreoffice` mode の出力保証

- `rows`, `table_candidates`, `print_areas`, `merged_cells` は既存の openpyxl / pandas ベース抽出をそのまま使う。
- `shapes` は LibreOffice UNO + OOXML drawing を使って best-effort で抽出する。
- `charts` は OOXML chart 定義 + LibreOffice UNO geometry で best-effort で抽出する。
- `auto_print_areas` は常に空とする。
- `colors_map`, `formulas_map`, cell hyperlink は既定では出さない。

### 5. LibreOffice 不在時 / 実行失敗時の挙動

- `mode="libreoffice"` 指定時に `soffice` または `uno` が使えない場合、処理は落とさず pre-analysis までの成果物で fallback workbook を返す。
- この fallback workbook には `rows`, `table_candidates`, `print_areas`, `merged_cells` を含め、`shapes` / `charts` は空配列にする。
- `PipelineState.fallback_reason` に `LIBREOFFICE_UNAVAILABLE` または `LIBREOFFICE_PIPELINE_FAILED` を設定する。
- `standard` / `verbose` は従来どおり COM 専用とし、LibreOffice への自動切り替えは行わない。

### 6. CLI / process API の互換性ルール

- `mode="libreoffice"` は抽出専用モードとして扱う。
- `process_excel(...)` / `ExStructEngine.process(...)` で `mode="libreoffice"` と以下の組み合わせを指定した場合、処理開始前に `ConfigError` を返す。
  - `pdf=True`
  - `image=True`
  - `auto_page_breaks_dir` を指定
- `extract_workbook(...)` / `ExStructEngine.extract(...)` で `include_auto_page_breaks=True` 相当になった場合も、`mode="libreoffice"` では `ConfigError` を返す。
- エラーは silent ignore ではなく hard fail に統一する。
- エラーメッセージは少なくとも以下の意図を含む明確な文言にする。
  - `libreoffice mode does not support PDF/PNG rendering`
  - `libreoffice mode does not support auto page-break export`
  - `use standard/verbose with Excel COM`

## モデル変更

### 1. Shape metadata

`BaseShape` に次の任意フィールドを追加する。

```python
class BaseShape(BaseModel):
    provenance: Literal["excel_com", "libreoffice_uno"] | None = None
    approximation_level: Literal["direct", "heuristic", "partial"] | None = None
    confidence: float | None = None
```

意味は以下。

- `provenance`: 抽出元 backend
- `approximation_level`:
  - `direct`: backend が直接持っている情報を採用
  - `heuristic`: 幾何推定などの推論を含む
  - `partial`: 一部は direct だが、一部欠落または代替手段を使った
- `confidence`: 0.0 から 1.0 の best-effort 信頼度

### 2. Chart metadata

`Chart` に同じ 3 フィールドを追加する。

```python
class Chart(BaseModel):
    provenance: Literal["excel_com", "libreoffice_uno"] | None = None
    approximation_level: Literal["direct", "heuristic", "partial"] | None = None
    confidence: float | None = None
```

既存の `name`, `chart_type`, `title`, `series`, `y_axis_title`, `y_axis_range`, `l`, `t`, `w`, `h`, `error` は維持する。

## 内部インターフェース仕様

### 1. pipeline 入力型

`ExtractionInputs.mode` は `libreoffice` を許容する。

```python
@dataclass(frozen=True)
class ExtractionInputs:
    file_path: Path
    mode: Literal["light", "libreoffice", "standard", "verbose"]
    ...
```

### 2. rich backend 抽象

shape / chart 抽出を backend 境界に寄せるため、内部専用 protocol を追加する。

```python
class RichBackend(Protocol):
    def extract_shapes(
        self,
        *,
        mode: Literal["libreoffice", "standard", "verbose"],
    ) -> dict[str, list[Shape | Arrow | SmartArt]]: ...

    def extract_charts(
        self,
        *,
        mode: Literal["libreoffice", "standard", "verbose"],
    ) -> dict[str, list[Chart]]: ...
```

- 既存 COM 抽出はこの protocol に合わせてラップする。
- 新規 `LibreOfficeRichBackend` を追加する。

### 3. LibreOffice session helper

LibreOffice UNO 呼び出しは subprocess 分離で扱う。

```python
@dataclass(frozen=True)
class LibreOfficeSessionConfig:
    soffice_path: Path
    startup_timeout_sec: float
    exec_timeout_sec: float
    profile_root: Path | None

class LibreOfficeSession:
    def __enter__(self) -> LibreOfficeSession: ...
    def __exit__(self, exc_type: object, exc: object, tb: object) -> None: ...
    def load_workbook(self, file_path: Path) -> object: ...
    def close_workbook(self, workbook: object) -> None: ...
    def extract_draw_page_shapes(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeDrawPageShape]]: ...
    def extract_chart_geometries(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeChartGeometry]]: ...
```

設定元環境変数:

- `EXSTRUCT_LIBREOFFICE_PATH`
- `EXSTRUCT_LIBREOFFICE_PYTHON_PATH`
- `EXSTRUCT_LIBREOFFICE_STARTUP_TIMEOUT_SEC`
- `EXSTRUCT_LIBREOFFICE_EXEC_TIMEOUT_SEC`
- `EXSTRUCT_LIBREOFFICE_PROFILE_ROOT`

### 4. OOXML helper

connector explicit ref と chart semantic を取るため、OOXML helper を追加する。

```python
@dataclass(frozen=True)
class DrawingShapeRef:
    drawing_id: int
    name: str
    kind: Literal["shape", "connector", "chart"]
    left: int | None
    top: int | None
    width: int | None
    height: int | None

@dataclass(frozen=True)
class DrawingConnectorRef:
    drawing_id: int
    start_drawing_id: int | None
    end_drawing_id: int | None

@dataclass(frozen=True)
class OoxmlChartInfo:
    name: str
    chart_type: str
    title: str | None
    y_axis_title: str
    y_axis_range: list[float]
    series: list[ChartSeries]
    anchor_left: int | None
    anchor_top: int | None
    anchor_width: int | None
    anchor_height: int | None
```

### 5. LibreOffice draw-page payload

`libreoffice` mode shape extraction adds a UNO draw-page payload model.

```python
@dataclass(frozen=True)
class LibreOfficeDrawPageShape:
    name: str
    shape_type: str | None = None
    text: str = ""
    left: int | None = None
    top: int | None = None
    width: int | None = None
    height: int | None = None
    rotation: float | None = None
    is_connector: bool = False
    start_shape_name: str | None = None
    end_shape_name: str | None = None
```

Connector resolution priority is fixed:
1. OOXML explicit ref (`stCxn`/`endCxn`)
2. UNO direct ref (`StartShape`/`EndShape`)
3. geometry heuristic (endpoint vs shape bbox)

`extract_shapes(mode="libreoffice")` uses the UNO draw-page payload as the
canonical emitted order when available. OOXML remains a supplemental source for
Excel-like shape type labels, connector arrowhead styles, explicit refs, and
heuristic endpoint geometry.

## 抽出アルゴリズム

### 1. shape / connector

`libreoffice` mode の shape / connector は以下の責務分担で組み立てる。

- UNO:
  - `DrawPage` から shape 一覧を取得
  - type, text, left/top, width/height, rotation を取得
  - `ConnectorShape` を識別
- OOXML drawing:
  - `xdr:sp`, `xdr:cxnSp`, `xdr:graphicFrame` を解析
  - `cNvPr id` と `stCxn/endCxn` を取得

node id と connector 解決のルールは固定する。

1. non-connector shape にだけシート内連番 `id` を振る
2. OOXML `cNvPr.id` と shape 名の両方を保持する
3. connector の begin/end は次の優先順で決める
   - OOXML `stCxn/endCxn` で解決できる場合:
     - `approximation_level="direct"`
     - `confidence=1.0`
   - UNO `StartShape/EndShape` が使える場合:
     - `approximation_level="direct"`
     - `confidence=0.9`
   - どちらも無い場合は幾何推定:
     - `approximation_level="heuristic"`
     - `confidence=0.6`
4. 幾何推定は connector の両端点と shape bbox の距離で nearest shape を選ぶ
5. 候補が見つからない側は `None` のままにする

### 2. chart

`libreoffice` mode の chart は以下の責務分担で組み立てる。

- OOXML / openpyxl:
  - chart 定義を読む
  - `chart_type`, `title`, `series`, `y_axis_title`, `y_axis_range` を構築
  - anchor から近似 geometry を得る
- UNO:
- `sheet.getCharts()` または `DrawPage` の `OLE2Shape` から chart geometry 候補を得る
  - v1 では LibreOffice 同梱 Python bridge subprocess から `sheet.getCharts()` と `DrawPage` の `OLE2Shape` を読む
  - `PersistName` と draw-page 順序を保持して OOXML chart との pairing 候補にする

pairing ルールは次のとおり。

1. OOXML chart の並び順を基準に 1 件ずつ構築する
2. UNO chart / OLE2Shape が同数で取得できる場合は順序で対応付ける
   - まず chart name / `PersistName` の一致を優先する
   - 残差だけを順序 pairing する
3. UNO geometry が無い場合は openpyxl anchor を使う
4. UNO geometry 使用時:
   - `approximation_level="partial"`
   - `confidence=0.8`
5. anchor のみ使用時:
   - `approximation_level="partial"`
   - `confidence=0.5`

## mode ごとの backend 解決

- `light`
  - rich backend 不使用
- `libreoffice`
  - pre-analysis: pandas / openpyxl
  - rich backend: LibreOffice UNO + OOXML
- `standard`
  - pre-analysis: 既存どおり
  - rich backend: Excel COM
- `verbose`
  - pre-analysis / rich backend とも既存どおり COM 前提

## テスト受け入れ条件

- API / CLI / MCP が `mode="libreoffice"` を受け付ける
- 無効 mode は従来どおり早期エラー
- `.xls` を `mode="libreoffice"` で指定すると早期エラー
- `process_excel(...)` / CLI で `mode="libreoffice"` と `pdf` / `image` / `auto_page_breaks_dir` を併用すると早期に `ConfigError` / 非ゼロ終了になる
- `extract_workbook(...)` / `ExStructEngine.extract(...)` で auto page-break を要求した場合も `mode="libreoffice"` では早期エラーになる
- `sample/flowchart/sample-shape-connector.xlsx` で connector の `begin_id/end_id` が十分数復元される
- `sample/basic/sample.xlsx` で chart が 1 件以上返り、title / series / geometry が埋まる
- LibreOffice 不在時は `rows` / `table_candidates` / `print_areas` / `merged_cells` を保った fallback を返す
- `standard` / `verbose` の既存 COM 系テストに回帰を出さない
- `model_dump(exclude_none=True)` により、新 metadata は未設定時に JSON に出ない

## 実装上の前提

- `soffice` と `uno` は optional dependency / 実行環境依存機能として扱う
- v1 では LibreOffice rendering は追加しないため、既存 `render` extra と切り離す
- 既存 sample を優先して回帰テストを作り、新規 fixture 追加は必要最小限にする

## 2026-03-07 LibreOffice bridge compatibility probe follow-up

### Issue

- system Python 自動検出が `import uno` と `PropertyValue` import の成功だけで候補を採用しており、bundled bridge script を実行できない Python 3.8 / 3.9 系も false positive として通してしまう。
- Debian 11 / Ubuntu 20.04 / その WSL で `python3-uno` が入っている場合、`mode="libreoffice"` は rich extraction 開始後に `_libreoffice_bridge.py` subprocess が `SyntaxError` で落ち、通常 fallback に戻ってしまう。

### Goal

- LibreOffice 用 Python 候補の互換性判定を「UNO import 可能」ではなく「bundled bridge をそのまま実行可能」まで引き上げる。
- 自動検出でも明示 override でも、bridge 非互換な Python は extraction 実行前に fail fast させる。

### Runtime contract

- LibreOffice bridge 用 Python 候補が「互換」と見なされる条件は次のすべてを満たすこと。
  - 実行ファイルが存在し、起動できる。
  - `uno` と `com.sun.star.beans.PropertyValue` を import できる。
  - `src/exstruct/core/_libreoffice_bridge.py` を専用 probe モードで実行し、timeout 内に exit code 0 で終了できる。
- `_resolve_python_path(...)` は、自動検出した system Python 候補のうち上記条件を満たすものだけを返す。
- `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` を指定した場合も、session 使用前に同じ probe を通す。非互換 override は rich extraction 実行中の遅延 `SyntaxError` ではなく、明確な runtime unavailable / incompatible error で失敗させる。
- 互換性判定は hard-coded な Python version 比較ではなく、bridge 実行 probe を唯一の真実とする。これにより、bridge 側の将来の構文・依存変更も probe で検知できる。

### Bridge probe contract

- `_libreoffice_bridge.py` に internal 用の `--probe` を追加する。
- `--probe` 実行時の bridge は以下を保証する。
  - module import と引数解析だけを通し、成功時は exit code 0 を返す。
  - `--host` / `--port` / `--file` を要求しない。
  - UNO socket 解決、document load、workbook access は行わない。
- runtime helper 側の probe はこの `--probe` を使う。`--help` や version string など副作用のある暗黙挙動には依存しない。

### Error handling

- probe 失敗条件は `OSError`、timeout、non-zero exit、bridge parse/runtime failure を含む。
- 自動検出候補がすべて probe 失敗した場合は、既存どおり「compatible Python runtime was not found」系の unavailable error に畳み込む。
- 明示 override が probe 失敗した場合は、設定済み Python path が bundled bridge と互換でないことが分かる文言にする。

### Verification

- bridge probe 成功時だけ system Python fallback が採用される unit test を追加する。
- `uno` import は通るが bridge probe が `SyntaxError` で失敗する候補を拒否する regression test を追加する。
- `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` 指定時も probe failure を fail-fast で surfacing する test を追加する。

## 2026-03-08 Linux LibreOffice CI smoke gate

### Goal

- `mode="libreoffice"` の rich extraction が Linux CI 上で実ランタイム smoke により継続的に保証されるようにする。
- probe unit test だけでなく、GitHub Actions 上で `soffice` 起動、UNO bridge、sample workbook 抽出までを必須ゲートにする。

### CI contract

- GitHub Actions に Linux 専用の required job `libreoffice-linux-smoke` を追加する。
- 当該 job は既存の unit matrix とは分離し、coverage upload の責務を持たない。
- runner は `ubuntu-24.04` に固定する。`ubuntu-latest` の image 切り替えに依存しない。
- runtime 準備として `libreoffice` と `python3-uno` を apt で導入する。
- smoke 実行時は `RUN_LIBREOFFICE_SMOKE=1` を設定し、`tests/core/test_libreoffice_smoke.py` を skip なしで実行する。
- smoke 失敗は fallback workbook 成功ではなく CI failure と見なし、PR merge を block する。

### Verification target

- `sample/flowchart/sample-shape-connector.xlsx` で connector の `begin_id/end_id` が復元される。
- `sample/basic/sample.xlsx` で chart title / series / geometry が取得される。
- `pytest.mark.libreoffice` の runtime gate は CI では disable せず、runtime unavailable の場合も job failure となる。

### Documentation

- README / README.ja / test requirements / task log に、Linux required smoke job の存在と実行条件を記載する。

## 2026-03-09 PR #76 additional review triage

### Accepted follow-up

#### LibreOffice smoke test should avoid backend-constant coupling

- `tests/core/test_libreoffice_smoke.py` は LibreOffice mode の end-to-end smoke を固定し、backend 実装定数そのものは固定しない。
- smoke では `chart.confidence == 0.8` のような exact 値比較を行わない。
- smoke は次を保証すればよい。
  - chart が返る
  - title / series / geometry が埋まる
  - `confidence` は `0.0 <= confidence <= 1.0` の範囲にある
- exact な `0.5 / 0.8` の confidence contract は、既存どおり backend 単体 test 側で担保する。

#### LibreOffice backend must not pay an extra probe-only session startup

- `LibreOfficeRichBackend.extract_shapes()` / `extract_charts()` は、実データ取得前に probe 専用の LibreOffice session を起動しない。
- runtime availability の確認は、`_read_draw_page_shapes()` または `_read_chart_geometries()` の実読みによって兼ねる。
- backend は probe-only `_runtime_checked` boolean に依存せず、実際に取得した cache の有無と各 read の成功/失敗で状態を表す。
- pipeline の `extract_shapes() -> extract_charts()` 経路では、LibreOffice session startup は最大 2 回までに抑える。
- ただし、shapes と charts の read は引き続き分離し、chart extraction failure 後も shape 成功結果を保持できる partial-success contract を壊さない。
- review comment の「二回目 failure が `_runtime_checked=True` で隠れる」は採用しない。read failure は従来どおり surfacing させる。

#### Test docstrings should remain grammatical and searchable

- `tests/core/test_mode_output.py` と `tests/cli/test_cli.py` の docstring は、意味の通る英文に揃える。
- `c l i` のような分割綴りは使わず `CLI` に統一する。
- 出力先や stdout 既定値を表す docstring は、対象関数名と期待結果が読める表現にする。

#### PR #76 should keep AGENTS.md changes scoped

- PR #76 は LibreOffice mode rollout を主題とするため、`AGENTS.md` の大規模な方針削除を同じ PR に混在させない。
- 現在の branch で削除された `AGENTS.md` の旧 section 2/3/4 は、この PR では restore する方針を優先する。
- もし `AGENTS.md` の整理自体を継続したい場合は、LibreOffice 変更とは分離した別 PR で扱う。

### Non-adopted review points

#### GitHub Actions package-manager consistency note

- `.github/workflows/pytest.yml` の `libreoffice-linux-smoke` は既存 `test` job と同じ `pip install -e .[...]` パターンを踏襲しており、新 job だけが不整合を持ち込んだわけではない。
- `defusedxml` 未導入という指摘は採用しない。core dependency なので editable install で導入される。
- `pytest-cov` 未使用は事実だが、動作不良ではなく cleanup レベルの論点として別件に留める。

#### MkDocs README navigation removal note

- `mkdocs.yml` / `docs/index.md` / `docs/README.en.md` / `docs/README.ja.md` の変更は docs build broken ではない。nav 参照と対象 file 削除が同期しているためである。
- この thread は機能バグとしては採用せず、必要なら「docs 導線再編をこの PR に残すか、別 PR に分離するか」の説明で解く。

## 2026-03-09 PR #76 Codacy command-injection follow-up

### Issue

- `python scripts/codacy_issues.py --pr 76 --min-level Warning` の 2026-03-09 時点の結果で、PR #76 に 1 件だけ `Security / Command Injection` が残っている。
- 指摘位置は `src/exstruct/core/libreoffice.py:825` で、`_run_bridge_probe_subprocess(...)` の `subprocess.run(...)` が対象になっている。
- 現行実装は `shell=False` と固定 argv を使っており、実害としての command injection というより、Semgrep/Codacy が `python_path` と `env` の trust boundary を追い切れていない false positive 寄りの状態である可能性が高い。

### Goal

- LibreOffice bridge probe の subprocess 呼び出しについて、Codacy が critical と判定しない形まで trust boundary を明確化する。
- 既存の runtime 互換性 probe 契約と、override/system candidate を fail-fast で reject する挙動は維持する。
- 修正スコープはまず probe helper に限定し、extraction/handshake の振る舞いは不用意に広げない。

### Probe subprocess contract

- `_run_bridge_probe_subprocess(...)` は、explicit な inherited `env=` を `subprocess.run(...)` に渡さない。
- probe subprocess の UTF-8 強制は `PYTHONIOENCODING` の環境変数注入ではなく、固定 argv 側の Python runtime option で行う。
- probe subprocess に渡す argv は次を満たす。
  - 実行ファイルは `_validated_runtime_path(...)` を通したローカル Python path
  - bridge script は repository 内の固定 `_libreoffice_bridge.py`
  - probe mode を示す fixed flag だけを追加する
- `_resolve_python_path(...)` / `_probe_libreoffice_bridge_failure(...)` は、override と system candidate の reject/accept 判定責務を引き続き持つ。Codacy 回避のために probe 自体を弱めない。
- 既存の `_build_subprocess_env(...)` allowlist は、少なくとも probe helper からは切り離す。必要なら extraction/handshake 専用 helper として残してよい。

### Fallback policy

- 上記の構造変更後も Codacy が同じ sink を `dangerous-subprocess-use-tainted-env-args` として報告する場合、その時点では static-analysis false positive と見なし、対象 call site に rule-specific suppression を最小範囲で追加する。
- suppression を入れる場合は、次をコードコメントで明記する。
  - `shell=False`
  - 実行ファイルと script path は local validated path
  - workbook path や command text を probe helper は受け取らない
  - したがって command injection ではなく analyzer limitation である

### Verification

- `tests/core/test_libreoffice_backend.py` に、probe helper が fixed argv で動き、explicit `env` を渡さないことを確認する regression test を追加する。
- 既存の `test_python_supports_libreoffice_bridge_filters_env` は、probe helper の新契約に合わせて置き換えるか、probe 用の「`env` を明示しない」検証へ更新する。
- 互換性判定の高レベル契約は維持する。
  - compatible runtime は probe success で accept される
  - incompatible override は fail-fast する
- 実装後は `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_libreoffice_bridge.py -q` と `uv run task precommit-run` を通す。
- push 後に `python scripts/codacy_issues.py --pr 76 --min-level Warning` を再実行し、当該 issue が消えていることを確認する。

## 2026-03-10 Windows LibreOffice CI smoke gate

### Goal

- `mode="libreoffice"` の hosted-runner smoke を Windows でも継続的に検証できるようにする。
- Linux smoke だけでは拾えない Windows 固有の `soffice.exe` path / process startup / UNO bridge 初期化の問題を CI で早期検出する。

### CI contract

- GitHub Actions に Windows 専用の smoke job `libreoffice-windows-smoke` を追加する。
- 当該 job は既存の unit matrix や COM test job と分離し、LibreOffice smoke だけを担当する。
- runner は `windows-2025` に固定する。
- runtime 準備として `choco install libreoffice-fresh -y --no-progress` を実行する。
- smoke 実行時は `RUN_LIBREOFFICE_SMOKE=1` と `FORCE_LIBREOFFICE_SMOKE=1` を設定する。
- ExStruct に `EXSTRUCT_LIBREOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe` を渡す。
- install 後は `soffice.exe` の存在確認と `--version` 実行で fail-fast し、その後 `tests/core/test_libreoffice_smoke.py -m libreoffice` を skip なしで実行する。
- Windows smoke は bundled Python の auto-detection をまず検証対象とし、`EXSTRUCT_LIBREOFFICE_PYTHON_PATH` は追加しない。

### Verification target

- Windows hosted runner 上で LibreOffice の起動コマンドが通る。
- `pytest.mark.libreoffice` smoke が runtime unavailable で skip されず、unavailable/incompatible 時は job failure になる。
- sample workbook の shape/chart extraction smoke が Windows 上でも green になる。

## 2026-03-10 Windows LibreOffice bundled Python auto-detection follow-up

### Issue

- `libreoffice-windows-smoke` で `soffice.exe --version` は通る一方、`tests/conftest.py::_has_libreoffice_runtime()` が `False` になり smoke test setup が `FORCE_LIBREOFFICE_SMOKE=1` で fail-fast した。
- 既存の `_resolve_python_path(...)` は `program/python.exe` のような直下候補しか見ておらず、Windows LibreOffice install の `python-core-*` 配下 layout を拾えない。

### Contract

- bundled LibreOffice Python の auto-detection は、従来の `program/python.exe` / `python.bin` / `python` 直下候補を維持する。
- 加えて Windows install で現れる `program/python-core-*/python.exe` と `program/python-core-*/bin/python.exe` 系候補も探索対象に含める。
- 既存どおり、採用条件は `_python_supports_libreoffice_bridge(...)` probe success とする。
- system Python fallback と explicit `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` override contract は変えない。

### Verification

- `tests/core/test_libreoffice_backend.py` に `python-core-*` 配下の `python.exe` が auto-detection で選ばれる regression test を追加する。
- `tests/test_conftest_libreoffice_runtime.py` と `python -m pre_commit run -a` を通し、runtime gate と型/lint を再確認する。

## 2026-03-10 Windows LibreOffice probe env follow-up

### Issue

- 最新の `libreoffice-windows-smoke` でも `tests/conftest.py::_has_libreoffice_runtime()` が `False` となり、`FORCE_LIBREOFFICE_SMOKE=1` により setup で fail-fast した。
- `resolve_python_path(...)` の互換性判定で使う bridge probe だけが allowlisted subprocess env を渡しておらず、Windows hosted runner の LibreOffice Python / UNO import に必要な runtime env を欠く可能性があった。

### Contract

- `_run_bridge_probe_subprocess(...)` も bridge handshake / extraction と同様に `_build_subprocess_env(...)` を使う。
- probe に渡す env は既存 allowlist に限定し、秘密値や無関係な env は forward しない。
- probe の argv contract (`python -X utf8 _libreoffice_bridge.py --probe`) は維持する。

### Verification

- `tests/core/test_libreoffice_backend.py` の probe subprocess regression test を、allowlisted env を forward する契約へ更新する。
- `tests/test_conftest_libreoffice_runtime.py` と `python -m pre_commit run -a` を再実行する。

## 2026-03-10 Windows LibreOffice workflow override follow-up

### Issue

- `libreoffice-windows-smoke` は probe env を forward した後も、Windows hosted runner 上で bundled Python auto-detection に依存する runtime gate のまま失敗し続けた。
- CI は `soffice.exe` の場所までは固定しているが、Chocolatey が導入する LibreOffice 26.x の bundled Python 配置は runner image / package variant に依存し得るため、workflow 側で明示 discovery した方が failure surface を狭められる。

### Contract

- Windows smoke workflow は LibreOffice install 後に bundled Python executable を探索し、`EXSTRUCT_LIBREOFFICE_PYTHON_PATH` として後続 step に引き渡す。
- discovery は既存 runtime helper が探索する bundled path 群 (`python.exe`, `python.bin`, `python`, `python-core-*`, `python-core-*\\bin`) に合わせる。
- bundled Python が見つからない場合は smoke test 実行前に job を fail し、program directory listing を残して原因調査を容易にする。

### Verification

- `.github/workflows/pytest.yml` が YAML として parse できることを確認する。
- Windows verify step で `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` の存在確認を行う。

## 2026-03-10 Windows LibreOffice bridge cwd follow-up

### Issue

- Windows smoke workflow で bundled `python.exe` の path discovery 自体は成功したが、`tests/conftest.py::_has_libreoffice_runtime()` は引き続き bridge probe を incompatible 扱いした。
- `_libreoffice_bridge.py` は module import 時点で `uno` を import するため、Windows の LibreOffice bundled Python は subprocess 実行時にも LibreOffice program directory を working directory として持つ必要がある。

### Contract

- LibreOffice bridge subprocess (`--probe`, `--handshake`, extraction) は `python_path` の親ディレクトリを `cwd` にして起動する。
- この `cwd` contract は allowlisted env と併用し、Windows bundled Python でも Linux system Python fallback でも同じ subprocess API で扱う。
- focused unit tests は probe / handshake / extraction subprocess が同じ `cwd` contract を使うことを検証する。

### Verification

- `uv run pytest tests/core/test_libreoffice_backend.py -q`
- `uv run pytest tests/test_conftest_libreoffice_runtime.py -q`
- `uv run task precommit-run`

## 2026-03-09 PR #76 latest review + Codacy re-triage

### Review-thread cleanup

- 追加レビューで重複して立った次の thread は、元の open thread に論点を集約する旨を返信して resolve 済みとする。
  - `discussion_r2904696477` -> combo chart series 指摘は `discussion_r2901508431` に集約
  - `discussion_r2904696479` -> connector direction rotation 指摘は `discussion_r2901508430` に集約
  - `discussion_r2901522451` -> redundant LibreOffice startup 指摘は `discussion_r2901509039` に集約
- 以後の実装 follow-up は、duplicate を除いた open thread だけを追えばよい。

### Accepted follow-up

#### Connector direction must honor OOXML rotation

- `src/exstruct/core/backends/libreoffice_backend.py::_resolve_direction()` は、`connector_info.direction_dx/direction_dy` をそのまま角度変換せず、`connector_info.rotation` を反映したベクトルを使って方位を決める。
- 実装は `_connector_endpoints()` と同じ `_rotate_connector_delta(...)` を再利用し、endpoint 推定と `direction` の幾何学的意味を一致させる。
- `dx == dy == 0` または OOXML delta 不足時は、従来どおり UNO box fallback を使う。
- regression test は「非回転だと従来どおり」「回転付き connector では endpoint と同じ向きへ方位が回る」の両方を固定する。

#### Combo-chart series extraction must scan every chart node

- `src/exstruct/core/ooxml_drawing.py::_extract_chart_series()` は、`plotArea` の最初の chart node だけで止めず、`_CHART_TAGS` に該当するすべての child node を document order で走査する。
- 各 chart node の `c:ser` を順に append し、combo chart の secondary series を欠落させない。
- `chart_type` 判定の「最初の chart node を代表値にする」契約はこの follow-up では変えない。今回は `series` 完全性だけを直す。
- regression test は `barChart + lineChart` の combo chart fixture を追加し、両 node の series が `name_range/x_range/y_range` 付きで保持されることを確認する。

#### OOXML connector arrowhead mapping must match begin/end semantics

- `src/exstruct/core/ooxml_drawing.py::_parse_connector_node()` の arrowhead mapping は、`a:headEnd -> begin_arrow_style`、`a:tailEnd -> end_arrow_style` に修正する。
- COM backend の `BeginArrowheadStyle / EndArrowheadStyle` と LibreOffice mode の意味を一致させる。
- 既存 test の expectation が誤実装に寄っている場合は更新し、head-only / tail-only の個別 regression test を追加して再発を防ぐ。

#### Bridge context resolution should guarantee at least one attempt

- `src/exstruct/core/_libreoffice_bridge.py::_resolve_context()` は、deadline 判定より先に `resolver.resolve(...)` を 1 回は試行する。
- ループ構造は「attempt -> failure なら deadline 判定 -> sleep/retry」とし、`timeout_sec <= 0` や極端に短い timeout でも no-attempt で `RuntimeError("Failed to resolve ...")` に落ちないようにする。
- timeout exhaustion時は、試行済みなら最後の UNO 例外を再送出する現在の意味を維持する。
- regression test は「最初の 1 回を必ず試す」「短い timeout でも no-attempt にならない」を直接確認する。

#### Test-docstring cleanup should cover the remaining generated cases

- 既に採用済みの `tests/core/test_mode_output.py` / `tests/cli/test_cli.py` に加え、`tests/core/test_pipeline.py` の不自然な自動生成 docstring も同じ sweep で直す。
- docstring は関数名の単純言い換えではなく、挙動と期待結果が読める英文に揃える。
- `CLI` の綴りは分割せず統一する。

#### Existing accepted follow-ups remain open

- `tests/core/test_libreoffice_smoke.py` の `chart.confidence == 0.8` を smoke 向け assertion へ緩める方針は継続する。
- `src/exstruct/core/backends/libreoffice_backend.py::_ensure_runtime()` の redundant startup 除去方針も継続する。
- `AGENTS.md` の PR scope 外削除をこの PR から外す方針も継続する。

### Codacy follow-up update

- push 後の `python scripts/codacy_issues.py --pr 76 --min-level Warning` では、残件は 1 件のままだが rule と行番号が変わった。
  - `Error | src/exstruct/core/libreoffice.py:806 | Semgrep_python.lang.security.audit.dangerous-subprocess-use-audit.dangerous-subprocess-use-audit | Security | Detected subprocess function 'run' without a static string.`
- 現在の対象は `_run_soffice_version_subprocess(...)` の `subprocess.run(...)` で、前回の `tainted-env-args` ではなく Semgrep の generic audit rule に切り替わっている。
- これは trusted local subprocess helper に対する構造的 false positive と判断する。
  - 実行ファイル path は `_validated_runtime_path(...)` を通す
  - `shell=False`
  - command text の組み立てなし
  - bridge helper では bundled local script を discrete argv で渡す
  - extract helper は workbook path を stdin 経由で渡す
- 次の対応は suppress-only ではなく、trusted subprocess sink であることを明記した narrow Semgrep suppression を helper 単位で付与する。
  - `_run_soffice_version_subprocess(...)`
  - `_run_bridge_probe_subprocess(...)`
  - `_run_bridge_extract_subprocess(...)`
  - `_run_bridge_handshake_subprocess(...)`
- 既存の `_spawn_trusted_subprocess(...)` に入っている `nosemgrep` と同じ rule id に揃え、helper comment では「local validated executable/script path」「shell=False」「no command-string assembly」を説明する。
- これにより、動作を変えずに static-analysis 境界だけを明示化する。

### Verification

- 追加実装後は少なくとも次を実行する。
  - `uv run pytest tests/core/test_libreoffice_backend.py tests/core/test_libreoffice_bridge.py tests/core/test_libreoffice_smoke.py tests/core/test_pipeline.py tests/core/test_mode_output.py tests/cli/test_cli.py -q`
  - `uv run task precommit-run`
- push 後に `python scripts/codacy_issues.py --pr 76 --min-level Warning` を再実行し、PR #76 の issue 数が減っていることを確認する。

## 2026-03-09 LibreOffice stderr cleanup masking on Windows

### Issue

- `mode="libreoffice"` の startup failure 後に `src/exstruct/core/libreoffice.py::_close_stderr_sink()` が一時 stderr log を `unlink()` すると、Windows で `PermissionError(13, "プロセスはファイルにアクセスできません。別のプロセスが使用中です。")` が出ることがある。
- この cleanup error が、本来 surfacing されるべき `LibreOfficeUnavailableError` を覆い隠し、pipeline 側では `libreoffice_pipeline_failed` として見えてしまう。

### Goal

- stderr sink cleanup を best-effort 化し、temporary log file の削除失敗が startup failure の真因を置き換えないようにする。
- cleanup の改善によって、本来の startup failure detail は引き続き error message に残す。

### Cleanup contract

- `_close_stderr_sink(stderr_sink, stderr_path)` は、sink close 後に stderr log file を best-effort で削除する。
- `unlink()` が `FileNotFoundError` の場合は従来どおり成功扱いとする。
- `unlink()` が `PermissionError` の場合は、Windows の handle release 遅延を吸収するため、短い bounded retry を行う。
- retry budget を超えても file が lock されたままなら、cleanup failure は黙って捨てる。temporary stderr log の残置は許容し、startup failure をマスクしてはならない。
- `_cleanup_failed_startup_process(...)` と `_start_soffice_startup_attempt(...)` は、stderr cleanup で `PermissionError` が起きても、元の `LibreOfficeUnavailableError` を `_LibreOfficeStartupAttemptError` として返す。

### Verification

- `tests/core/test_libreoffice_backend.py` に、`_close_stderr_sink()` が一時的な `PermissionError` を retry 後に解消できる regression test を追加する。
- 同 test file に、stderr log unlink が lock され続けても `_start_soffice_startup_attempt(...)` が `PermissionError` ではなく startup failure を返す regression test を追加する。

## 2026-03-10 PR79 LibreOffice Windows smoke runtime gate hardening

### Background

- Windows CI (`libreoffice-windows-smoke`) can have transiently slow `soffice --version` startup right after Chocolatey install.
- `tests/conftest.py::_has_libreoffice_runtime()` currently treats any timeout from the version probe as runtime unavailable, which can create a false negative and trip `FORCE_LIBREOFFICE_SMOKE=1`.

### Spec

- `_has_libreoffice_runtime()` keeps strict checks for:
  - `soffice` executable path existence
  - bridge-compatible Python resolution
- Version probe policy is refined:
  - On `subprocess.TimeoutExpired` from `soffice --version`, do **not** immediately mark runtime unavailable.
  - Attempt one fallback runtime viability check by launching `LibreOfficeSession.from_env()` and entering/exiting the session.
  - If fallback session startup succeeds, treat runtime as available (`True`).
  - If fallback session startup fails with `LibreOfficeUnavailableError`, treat runtime as unavailable (`False`).
  - Unexpected exceptions from the fallback remain loud (raise) to avoid masking regressions.
- Existing behavior for `FileNotFoundError`, `OSError`, and `CalledProcessError` in version probe remains `False`.

### Function contracts

- `_has_libreoffice_runtime() -> bool`
  - Inputs: none (uses env + runtime helpers)
  - Output: runtime availability decision with timeout fallback behavior above
  - Side effect: may create a short-lived `LibreOfficeSession` during timeout fallback path


## 2026-03-10 PR79 follow-up: slow Windows version probe retry

### Background

- Previous fix added session fallback when `soffice --version` times out, but CI still fails on Windows with `FORCE_LIBREOFFICE_SMOKE=1`.
- The failure window suggests first-run startup latency can exceed 5 seconds and may also cause immediate session startup to miss the 15-second budget under cold install conditions.

### Spec

- In `tests/conftest.py::_has_libreoffice_runtime()`:
  - Keep the fast `soffice --version` probe (5s) for normal cases.
  - If that first probe times out, retry `soffice --version` once with an extended timeout (30s).
  - If the extended retry succeeds, return `True` and skip session fallback.
  - If the extended retry still times out, keep the existing session fallback probe (`LibreOfficeSession.from_env()`).
  - `FileNotFoundError`, `OSError`, `CalledProcessError` from either version probe still map to `False`.

### Function contracts

- `_has_libreoffice_runtime() -> bool`
  - Performs at most two version probes before session fallback:
    1. timeout=5.0
    2. timeout=30.0 (only after first timeout)
  - Returns `True` when retry probe or fallback session succeeds.
  - Returns `False` on expected runtime-unavailable failures.


## 2026-03-11 PR #79 Windows LibreOffice smoke stabilization

- `_validated_runtime_path(path: Path) -> Path`
  - Windows で `soffice.exe` が指定された場合、同ディレクトリの `soffice.com` が存在すれば `soffice.com` に正規化する。
  - 非 Windows では入力 path を維持する。
- `_which_soffice() -> Path | None`
  - PATH 探索順を `soffice` → `soffice.com` → `soffice.exe` とし、見つかった path は runtime path 正規化を通す。
- GitHub Actions `libreoffice-windows-smoke`
  - `EXSTRUCT_LIBREOFFICE_PATH` は `soffice.com` を優先し、存在しない場合のみ `soffice.exe` を fallback とする。
  - Verify step は `--version` 実行後に `$LASTEXITCODE` を検証し、非ゼロを fail-fast する。
