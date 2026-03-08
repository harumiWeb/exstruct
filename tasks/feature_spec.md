# Feature Spec

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
