# Feature Spec

## 2026-04-22 README English/Japanese parity refresh

### Goal

- Bring `README.md` back in line with the heavily edited `README.ja.md`.
- Remove English-only sections or details that no longer exist in the Japanese README.
- Preserve the same public structure, examples, and interface positioning across both README files while keeping the English text idiomatic.

### Public contract summary

- `README.md` should describe the same interfaces, quick starts, examples, and support notes as `README.ja.md`.
- The English README should not retain extra positioning guidance, extra MCP operational notes, or longer example explanations that the Japanese README no longer ships.
- This task is documentation parity only; no code, CLI behavior, API behavior, or ADR policy changes are introduced.

### Permanent destinations

- `README.md`
  - Updated English public-facing project overview, quick starts, examples, and reference links.
- `README.ja.md`
  - Remains the parity source for this specific cleanup pass; no content changes required in this task.
- No additional `dev-docs/` or `docs/` migration is required because this change only re-syncs an already-public README.

### Verification

- `git diff --check -- README.md tasks/feature_spec.md tasks/todo.md`

### ADR verdict

- `not-needed`
- rationale: this is a documentation parity refresh for an existing public README, not a new design or policy decision.

## 2026-04-22 light-mode print areas / OOXML drawing resilience

### Goal

- Align `light` mode print-area behavior across `extract`, `process_excel`, CLI, and engine export so the accepted ADR/docs contract is consistent on every public path.
- Limit OOXML drawing parse failures to the malformed worksheet so healthy sheets keep their best-effort shapes/charts.

### Public contract summary

- `mode="light"` keeps `print_areas` in default structured output and allows `print_areas_dir` side-output generation on `process_excel` and CLI paths.
- `FilterOptions.include_print_areas=None` means automatic inclusion for all modes; callers must pass `False` explicitly to suppress print areas.
- OOXML rich extraction remains best-effort, but a malformed drawing part on one sheet must not erase healthy rich artifacts from other sheets in the same workbook.

### Permanent destinations

- `dev-docs/specs/excel-extraction.md`, `docs/api.md`, `docs/cli.md`, and `ADR-0010` already hold the durable mode contract; no new permanent spec is needed.
- `dev-docs/testing/test-requirements.md` must reflect the corrected `light` print-area default.
- `docs/generated/models.md` must be regenerated because the `FilterOptions.include_print_areas` description changes.

### Verification

- `uv run pytest tests/engine/test_engine.py tests/core/test_mode_output.py tests/cli/test_cli.py tests/core/test_ooxml_drawing.py -q`
- `uv run python scripts/gen_model_docs.py`
- `uv run task precommit-run`

### ADR verdict

- `not-needed`
- rationale: this work restores the accepted `ADR-0010` contract and fallback behavior; it does not introduce a new policy decision.

## 2026-04-22 PR #129 review follow-up

### Goal

- Address the actionable PR review feedback for the light-mode OOXML rich baseline without changing the accepted mode contract.
- Keep `light` mode resilient to unexpected OOXML rich-extraction failures and reduce worksheet-metrics overhead in the new OOXML drawing path.

### Public contract summary

- `process_excel()` should continue to inherit the engine default `include_print_areas=None` behavior rather than hard-coding print-area inclusion.
- `light` mode still supports best-effort OOXML shapes/charts, but unexpected rich-extraction failures must degrade to the existing cells/tables fallback instead of aborting extraction.
- Architecture docs must describe `OoxmlRichBackend` as the concrete `RichBackend` for `light`.

### Permanent destinations

- `src/exstruct/__init__.py` should keep `process_excel()` aligned with the engine auto-filter contract.
- `src/exstruct/core/pipeline.py` and `src/exstruct/core/ooxml_drawing.py` should hold the resilience and performance fixes for light-mode OOXML enrichment.
- `dev-docs/architecture/pipeline.md` should reflect the current backend set and light-mode routing.
- No new ADR/spec document is needed; this is implementation/doc follow-up within the accepted `ADR-0010` direction.

### Verification

- `uv run pytest tests/core/test_pipeline.py tests/core/test_mode_output.py tests/core/test_ooxml_drawing.py -q`
- `uv run task precommit-run`

### ADR verdict

- `not-needed`
- rationale: this is review-driven hardening and architecture-doc alignment for an already accepted design.

## 2026-04-22 PR #129 review follow-up (second pass)

### Goal

- Address the next review pass for PR `#129`, focusing on one remaining OOXML per-sheet resilience hole and stale public/docs wording.
- Normalize the specific YAML/Markdown files flagged for CRLF line endings so repo tooling and review bots stop reporting newline-only issues.

### Public contract summary

- README examples and non-COM fallback wording must describe the current `light` / OOXML-rich contract accurately.
- A corrupt OOXML zip member for one worksheet drawing must still be handled at the per-sheet boundary when possible.
- Line-ending-only cleanup must not change behavior; it only restores LF normalization for the flagged files.

### Permanent destinations

- `README.md` and `README.ja.md` should keep the public extraction-mode wording aligned with `ADR-0010` and the current implementation.
- `src/exstruct/core/ooxml_drawing.py` should keep the per-sheet OOXML drawing error boundary as narrow as safely possible.
- `.agents/skills/exstruct-cli/agents/openai.yaml`, `dev-docs/agents/coding-guidelines.md`, and `mkdocs.yml` should be normalized back to LF.

### Verification

- `uv run pytest tests/core/test_ooxml_drawing.py -q`
- `uv run task precommit-run`

### ADR verdict

- `not-needed`
- rationale: this is review-driven contract wording cleanup and newline normalization on top of the existing accepted design.

## 2026-03-19 v0.7.0 release closeout

### Goal

- Publish the `v0.7.0` release-prep artifacts for the workbook editing work delivered through issue `#99`.
- Collapse temporary issue and review logs after moving the only remaining maintainer-facing rule into permanent documentation.
- Keep a compact closeout record that states where permanent information now lives and how the release-prep work was verified.

### Permanent destinations

- `CHANGELOG.md`
  - Holds the `0.7.0` `Added` / `Changed` / `Fixed` summary for the public release.
- `docs/`
  - `docs/release-notes/v0.7.0.md` records the user-facing release narrative for issue `#99`, review follow-ups, and maintainer-facing documentation work.
  - `mkdocs.yml` keeps the canonical `Release Notes` navigation entry for `v0.7.0`.
- `dev-docs/architecture/overview.md`
  - Records the legacy monkeypatch compatibility precedence note for compatibility shims.
- `dev-docs/specs/`
  - No new spec migration was required for this closeout; the existing editing API and CLI specs remain the canonical contract source.
- `dev-docs/adr/`
  - No new ADR or ADR update was required; existing `ADR-0006` and `ADR-0007` continue to cover the policy boundary.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain only this release-closeout record plus verification, not the historical phase-by-phase work log.

### Constraints

- Public API / CLI / MCP contracts and backend selection policy remain unchanged in this closeout.
- `README.md` and `docs/index.md` do not gain direct release-note links; `mkdocs.yml` stays the canonical route.
- `uv.lock` is not fully regenerated; only the local `exstruct` package version is aligned to `0.7.0`.

### Verification

- `uv run task build-docs`
- `uv run task precommit-run`
- `rg -n "0\.7\.0|v0\.7\.0" pyproject.toml uv.lock CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.0.md`
- `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
- `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.0.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md dev-docs/architecture/overview.md`

### ADR verdict

- `not-needed`
- rationale: this was release preparation and task-log retention cleanup. The policy decisions already live in `ADR-0006`, `ADR-0007`, and the editing specs.

## 2026-03-21 v0.7.1 release closeout

### Goal

- Publish the `v0.7.1` release-prep artifacts for the CLI and package import startup optimization work delivered through issues `#107`, `#108`, and `#109`.
- Collapse the temporary issue and review logs for `#107` and `#108` after confirming that the durable contract and design rationale already live in permanent documentation.
- Keep a compact closeout record that states where permanent information now lives and how the release-prep work was verified.

### Public contract summary

- `--auto-page-breaks-dir` is always listed in extraction CLI help output and validated only when the flag is requested at runtime.
- `exstruct --help`, `exstruct ops list`, non-edit CLI routing, `import exstruct`, and `import exstruct.engine` now defer heavy imports until execution actually needs them.
- Public exported symbol names from `exstruct` and `exstruct.edit` remain stable; only import timing changed.
- The edit CLI `validate` subcommand keeps its narrow historical error boundary and must still propagate `RuntimeError`.
- No new CLI commands, MCP payload shapes, or backend-selection policy changes are introduced in this closeout.

### Permanent destinations

- `CHANGELOG.md`
  - Holds the `0.7.1` `Added` / `Changed` / `Fixed` summary for the public release.
- `docs/`
  - `docs/release-notes/v0.7.1.md` records the user-facing release narrative for issues `#107`, `#108`, and `#109`.
  - `mkdocs.yml` keeps the canonical `Release Notes` navigation entry for `v0.7.1`.
  - `docs/cli.md` remains the canonical public CLI contract for extraction help/runtime validation behavior.
- `README.md` and `README.ja.md`
  - Retain the public-facing wording for extraction runtime validation and CLI behavior that shipped with issue `#107`.
- `dev-docs/specs/`
  - `dev-docs/specs/excel-extraction.md` remains the canonical internal guarantee for extraction CLI runtime validation.
- `dev-docs/architecture/overview.md`
  - Records the durable lightweight-startup rule for package `__init__` files, CLI routing, and `exstruct.engine`.
- `dev-docs/adr/`
  - `ADR-0008` remains the canonical policy source for runtime capability validation in the extraction CLI.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain only the release-closeout records plus verification, not the detailed issue-by-issue implementation log.

### Constraints

- `README.md` and `docs/index.md` do not gain direct release-note links; `mkdocs.yml` stays the canonical navigation route.
- `uv.lock` is not fully regenerated; only the editable `exstruct` package version is aligned to `0.7.1`.
- This closeout does not add a new ADR or new permanent spec document; it only points to the existing permanent sources for the shipped behavior.

### Verification

- `uv run pytest tests/cli/test_cli.py tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`
- `uv run task build-docs`
- `uv run task precommit-run`
- `rg -n "0\.7\.1|v0\.7\.1" CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.1.md`
- `rg -n '^version = "0\.7\.1"$' pyproject.toml uv.lock`
- `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
- `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.1.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md`

### ADR verdict

- `not-needed`
- rationale: this was release preparation and task-log retention cleanup. The shipped policy decisions already live in `ADR-0008`, the extraction docs/specs, and the architecture note.

## 2026-03-21 issue #115 ExStruct CLI Skill

### Goal

- Add one installable Skill that teaches AI agents how to use the existing ExStruct editing CLI safely and consistently.
- Keep the current CLI, Python API, and MCP runtime contracts unchanged while packaging operational knowledge into repo-owned skill assets and durable documentation.
- Record the long-lived policy choices behind the Skill structure, CLI/MCP boundary, and distribution model in permanent documents before implementation logs are compressed.

### Public contract summary

- No changes to `exstruct patch`, `exstruct make`, `exstruct ops`, `exstruct validate`, `exstruct.edit`, or MCP tool payloads.
- Public docs gain one new documented interface layer: a single installable Skill named `exstruct-cli` for agent-side CLI workflows.
- `README.md` and `README.ja.md` must describe:
  - where the Skill source lives in the repository
  - how to install/copy it into an agent runtime
  - when to use the Skill versus MCP guidance
  - minimal usage examples / prompts

### Internal Skill contract

- Canonical repo source: `.agents/skills/exstruct-cli/`
- Required files:
  - `SKILL.md`
  - `agents/openai.yaml`
  - `references/command-selection.md`
  - `references/safe-editing.md`
  - `references/ops-guidance.md`
  - `references/verify-workflows.md`
  - `references/backend-constraints.md`
- `SKILL.md` must stay lean and contain only trigger-oriented frontmatter, command selection rules, safety rules, workflow steps, and direct links to `references/`.
- `references/` must hold the detailed operational knowledge and avoid duplicating long catalogs inside `SKILL.md`.
- `scripts/` and `assets/` are out of scope for the initial implementation unless a concrete deterministic need appears during writing or validation.

### Permanent destinations

- `.agents/skills/exstruct-cli/`
  - Canonical source for the installable Skill assets.
- `README.md` and `README.ja.md`
  - Public installation and usage guidance for the Skill.
- `dev-docs/specs/`
  - `dev-docs/specs/exstruct-cli-skill.md` is the canonical internal spec for the ExStruct CLI Skill contract and required reference structure.
- `dev-docs/adr/`
  - `ADR-0009-single-cli-skill-for-agent-workflows.md` records the policy-level decisions: single Skill, repo-owned source of truth, and CLI-versus-MCP documentation boundary.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Keep the working record only until the durable information above is written.

### Constraints

- Use the existing repository convention for skills (`.agents/skills/` + `agents/openai.yaml`); do not add `.claude/skills/` mirrors or sync automation in this issue.
- Keep the Skill focused on local CLI workflows; MCP host-policy and transport concerns stay documented separately under the existing MCP docs.
- Follow docs parity rules: any public README change in English must be mirrored in `README.ja.md` in the same change.
- Reuse the external `skill-creator` helper scripts for `openai.yaml` generation and lightweight validation rather than adding new repo-local validation code for this issue.

### Verification

- `python <skill-creator>/scripts/generate_openai_yaml.py .agents/skills/exstruct-cli --interface ...`
- `python <skill-creator>/scripts/quick_validate.py .agents/skills/exstruct-cli`
- `uv run task precommit-run`
- Manual scenario review of:
  - create-vs-edit command selection
  - `validate -> dry-run -> inspect -> apply -> verify` guidance
  - unsupported-op / backend-constraint handling
  - CLI-vs-MCP routing guidance
  - README English/Japanese parity

### ADR verdict

- `recommended`
- rationale: the change turns AI-agent operational workflow into a durable repository rule and resolves recurring tradeoffs around single-skill packaging, repo source of truth, and the CLI-versus-MCP boundary.

## 2026-04-16 SECURITY.md policy

### Goal

- Add a root-level `SECURITY.md` that GitHub can recognize as the repository security policy.
- Direct security reports to `harumiweb.security@gmail.com` and keep sensitive disclosures out of public issue threads when they are not already public.
- Keep the change documentation-only with no code, package, CLI, MCP, or MkDocs navigation impact.

### Public contract summary

- The repository gains one new public policy document: `SECURITY.md`.
- Supported versions are defined as the latest release only.
- Security vulnerabilities should be reported by email first.
- Public GitHub issues remain appropriate for non-security problems and already-public, non-sensitive discussion.

### Permanent destinations

- `SECURITY.md`
  - Canonical public security policy document for responsible disclosure and supported-version guidance.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain only this compact implementation record and verification evidence for the session.

### Constraints

- `SECURITY.md` is English-only for this change.
- `README.md`, `README.ja.md`, `docs/`, and `mkdocs.yml` remain unchanged.
- The supported-version policy must avoid hard-coding a specific release number and instead describe support as "latest release".

### Verification

- `rg -n "Security Policy|harumiweb.security@gmail.com|Latest release|GitHub Issues" SECURITY.md`
- `git diff --check -- SECURITY.md tasks/feature_spec.md tasks/todo.md`
- `uv run task precommit-run`
- `uv run pytest -q`

### ADR verdict

- `not-needed`
- rationale: this adds a single public repository policy document without changing architecture, public API design, or long-lived internal tradeoff policy.

## 2026-04-16 issue #77 LibreOffice typed workbook handle

### Goal

- Replace the raw `dict` token returned by `LibreOfficeSession.load_workbook()` with a typed workbook handle.
- Give `LibreOfficeSession.close_workbook()` meaningful session-local cleanup instead of a no-op.
- Keep the current LibreOffice extraction mode, fallback behavior, and bridge subprocess lifecycle unchanged.

### Contract summary

- `LibreOfficeSession.load_workbook()` returns a frozen typed handle that stores the resolved workbook path and the owning session identity.
- `LibreOfficeSession.close_workbook()` validates that the handle belongs to the current session, rejects rehydrated handles whose `file_path` no longer matches the registered workbook id, becomes idempotent for repeated close attempts, and clears any session-local bridge cache entries for that workbook.
- `LibreOfficeSession.extract_draw_page_shapes()` and `extract_chart_geometries()` continue to support path-based extraction, but may also consume the typed workbook handle so callers can follow a typed lifecycle.
- `LibreOfficeRichBackend.session_factory` accepts the structural rich-extraction session contract, including legacy path-only sessions and lifecycle-aware sessions, rather than only the concrete built-in `LibreOfficeSession`.
- No public CLI, MCP, extraction-mode, fallback, or serialization contracts change in this issue.

### Permanent destinations

- `src/exstruct/core/libreoffice.py`
  - Canonical implementation for the typed LibreOffice workbook handle and close semantics.
- `src/exstruct/core/backends/libreoffice_backend.py`
  - Updated to consume the typed session lifecycle without changing backend policy.
- `tests/core/test_libreoffice_backend.py`
  - Regression coverage for typed handle behavior, ownership checks, idempotent close, and cache invalidation.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain the compact planning and verification record for this issue.

### Constraints

- Do not change the bridge subprocess contract in `src/exstruct/core/_libreoffice_bridge.py`; workbook documents are still opened and closed per bridge invocation.
- Do not change backend fallback policy or session startup/shutdown behavior.
- Keep backward compatibility for current path-based extraction helpers while introducing the typed handle.

### Verification

- `uv run pytest tests/core/test_libreoffice_backend.py -q`
- `uv run task precommit-run`

### ADR verdict

- `not-needed`
- rationale: this is an internal contract hardening change that preserves existing extraction policy and runtime behavior; the durable rationale can stay in the task record.

## 2026-04-21 Pure-Python rich extraction for light-mode environments

### Goal

- Provide shapes/connectors/charts from pure-Python OOXML parsing on `.xlsx/.xlsm` even in environments that currently only get `light`-level extraction.
- Treat this as strengthening the non-COM/no-LibreOffice environment, not as removing or redefining the current LibreOffice runtime path.
- Record the now-chosen policy that the richer pure-Python artifacts will be exposed by strengthening `light` itself.

### Investigation summary

- `src/exstruct/core/ooxml_drawing.py` already parses OOXML worksheet drawings into `SheetDrawingData` with:
  - shapes (`OoxmlShapeInfo`)
  - connectors (`OoxmlConnectorInfo`)
  - charts (`OoxmlChartInfo`)
- `src/exstruct/core/backends/libreoffice_backend.py` already contains a pure-OOXML path:
  - `_build_shapes_from_ooxml(...)` emits shapes/arrows without UNO snapshots
  - `extract_charts(...)` already builds chart metadata from OOXML and uses LibreOffice only to refine geometry/confidence
- The current blocker is architectural, not fundamental parsing capability:
  - `extract_shapes()` always calls `_read_draw_page_shapes()`
  - `extract_charts()` always calls `_read_chart_geometries()`
  - if the LibreOffice session is unavailable, the whole rich path falls back before the existing OOXML-only builders can be used
- Current OOXML geometry still needs hardening for Python-only use:
  - `_marker_to_points()` uses fixed default column/row sizes
  - `sample/flowchart/sample-shape-connector.xlsx` uses custom worksheet column widths, so pure anchor-based placement will drift unless sheet metrics are parsed
  - many shapes/connectors also carry `a:xfrm` offsets/extents, which can serve as a higher-quality pure-Python geometry source than the current default anchor approximation
- Chart extraction is closest to ready for Python-only use:
  - OOXML already provides chart type, title, series, axis range, and anchor geometry
  - LibreOffice chart geometry is an optional refinement layer today, not the primary metadata source
- A public-contract question remains for backend metadata:
  - `src/exstruct/models/__init__.py` and `schemas/*.json` currently allow only `excel_com` and `libreoffice_uno` in `provenance`
  - a true Python-only path should either introduce a new provenance label or explicitly decide that `libreoffice` mode provenance remains mode-oriented rather than runtime-oriented

### Proposed contract direction

- The user-selected contract direction is to strengthen `light` itself rather than add another non-COM rich mode.
- Environment capability and mode semantics are therefore intentionally aligned:
  - non-COM/non-LibreOffice hosts gain pure-Python shapes/connectors/charts on `.xlsx/.xlsm`
  - `light` becomes the public entrypoint for that baseline capability
- Current codebase evidence says changing `light` directly is a public-contract change:
  - [docs/cli.md](/mnt/c/dev/python/exstruct/docs/cli.md:114) defines `light` as cells + table candidates + print areas only.
  - [dev-docs/specs/excel-extraction.md](/mnt/c/dev/python/exstruct/dev-docs/specs/excel-extraction.md:19) defines `light` as cells/tables only, no rich artifacts.
  - [dev-docs/testing/test-requirements.md](/mnt/c/dev/python/exstruct/dev-docs/testing/test-requirements.md:203) locks `MODE-02` to shapes/charts empty in `light`.
  - [dev-docs/adr/ADR-0001-extraction-mode-boundaries.md](/mnt/c/dev/python/exstruct/dev-docs/adr/ADR-0001-extraction-mode-boundaries.md:15) treats mode responsibility boundaries as an explicit decision.
- The chosen policy path is therefore:
  - keep building the pure-Python rich backend first
  - expose it through `light`
  - update the mode contract, docs, schemas if needed, and tests in the same change
- Preserve the existing fallback policy from `ADR-0002`:
  - when the pure-Python rich path cannot safely extract rich artifacts, fall back to cells + table candidates + pre-com artifacts
- Keep `.xls` out of scope for this phase; this work applies only to OOXML workbooks.
- Keep SmartArt and grouped-shape reconstruction out of the first phase unless a specific fixture proves they are required for the first usable release.

### Staged implementation plan

1. Build a pure-Python rich extraction path independent of LibreOffice runtime.
   - Reuse [ooxml_drawing.py](/mnt/c/dev/python/exstruct/src/exstruct/core/ooxml_drawing.py:126) as the canonical source for OOXML shapes/connectors/charts.
   - Avoid tying the baseline implementation to [libreoffice_backend.py](/mnt/c/dev/python/exstruct/src/exstruct/core/backends/libreoffice_backend.py:73) naming or runtime assumptions more than necessary.
2. Improve pure-Python geometry fidelity in `src/exstruct/core/ooxml_drawing.py`.
   - Parse sheet row heights and column widths from worksheet XML instead of relying on `_DEFAULT_COLUMN_WIDTH_POINTS` / `_DEFAULT_ROW_HEIGHT_POINTS` alone.
   - Prefer child `a:xfrm` left/top/width/height for shapes/connectors when present and valid; use anchor geometry as the fallback or for chart frames with zero transforms.
   - Add focused tests for custom-width/custom-height sheets so geometry does not regress silently.
3. Apply the chosen exposure path in the pipeline and public entrypoints.
   - `light` should surface the new pure-Python rich baseline for `.xlsx/.xlsm`.
   - `libreoffice` should remain the optional enrichment path above that baseline.
   - Update `MODE-02` expectations explicitly instead of letting the old empty-shape assumption linger.
4. Lock down Python-only connector resolution and chart extraction with tests.
   - Add tests that assert emitted shapes/connectors/charts without any LibreOffice session snapshots.
   - Add regression tests for mixed cases: OOXML-only and OOXML + LibreOffice enrichment.
5. Resolve metadata/documentation policy before merging behavior changes.
   - Decide whether to add a new `provenance` literal such as `python_ooxml` / `ooxml`.
   - Update `dev-docs/specs/excel-extraction.md`, `dev-docs/specs/data-model.md`, `docs/cli.md`, `docs/api.md`, and `docs/mcp.md` once the final exposure path is chosen.

### Constraints

- Do not weaken `ADR-0002` fallback behavior; the change should improve rich extraction availability, not change the safe fallback shape on true failure.
- Do not route `.xls` through a partially supported pseudo-OOXML path.
- Do not introduce Excel/LibreOffice-specific logic into `pipeline.py`; backend composition stays inside the rich backend layer.
- Treat any new provenance literal as a public serialization/schema change that requires coordinated model/docs/schema updates.
- If `light` is changed to include rich artifacts, treat that as an explicit mode-boundary change rather than an internal optimization.

### Implementation status

- Completed in the first implementation pass:
  - added `src/exstruct/core/backends/ooxml_backend.py` as the pure-Python OOXML rich backend for `.xlsx/.xlsm`
  - wired `light` to use that backend and preserve emitted shapes/charts in workbook assembly
  - changed `libreoffice` fallback handling so the OOXML baseline is preserved when LibreOffice runtime enrichment is unavailable
- Completed in the geometry-fidelity follow-up:
  - `src/exstruct/core/ooxml_drawing.py` now parses worksheet row heights and column widths into `SheetDrawingMetrics`, and anchor fallback uses those metrics instead of fixed defaults alone
  - shapes and connectors now prefer `a:xfrm` absolute left/top when the transform also carries a non-zero size, which fixes large placement drift on `sample/flowchart/sample-shape-connector.xlsx`
  - charts keep the same safe fallback behavior: if `graphicFrame/xfrm` is a zero placeholder, anchor geometry remains the source of placement and size
  - regression coverage now includes dedicated `tests/core/test_ooxml_drawing.py` cases for custom metrics, two-cell anchors, and transform-preferred fixture placement
  - introduced `python_ooxml` as a `provenance` literal in models and generated schemas
  - added regression tests for:
    - light-mode OOXML rich extraction without COM
    - workbook assembly retaining rich artifacts in light mode
    - LibreOffice-unavailable fallback preserving the OOXML baseline
    - light-mode raw-data collection retaining charts
- Completed in the durable-contract follow-up:
  - `ADR-0010` is accepted and supersedes `ADR-0001`, which is now marked `superseded`
  - internal specs now describe `light` as the pure-Python OOXML-rich baseline and document `python_ooxml` as a public backend-metadata provenance value
  - public docs now describe `light` as the preferred non-COM baseline for `.xlsx/.xlsm` and `libreoffice` as the optional enrichment path above it

### Permanent destinations

- `dev-docs/adr/`
  - `dev-docs/adr/ADR-0010-light-mode-as-the-pure-python-rich-ooxml-baseline.md` now holds the accepted policy decision that changes `light` into the pure-Python OOXML-rich baseline for `.xlsx/.xlsm`.
- `dev-docs/specs/excel-extraction.md`
  - Canonical internal spec for the updated non-COM rich extraction contract.
- `dev-docs/specs/data-model.md` and `schemas/`
  - Canonical internal destination for backend metadata literals and serialization semantics.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Temporary implementation record now that the ADR/spec updates have been authored.

### ADR verdict

- `required`
- rationale:
  - this change alters the `light` mode boundary for non-COM rich extraction
  - it may also change the meaning of backend metadata (`provenance`) exposed through public models and schemas
- affected domains:
  - extraction mode
  - backend fallback
  - backend metadata / serialization
- existing ADR candidates:
  - `dev-docs/adr/ADR-0001-extraction-mode-boundaries.md`
  - `dev-docs/adr/ADR-0002-rich-backend-fallback-policy.md`
- suggested next action:
  - `new-adr` completed via `ADR-0010` draft
- evidence triad:
  - specs:
    - `dev-docs/specs/excel-extraction.md`
    - `dev-docs/specs/data-model.md`
    - `docs/cli.md`
    - `docs/mcp.md`
  - src:
    - `src/exstruct/core/backends/libreoffice_backend.py`
    - `src/exstruct/core/ooxml_drawing.py`
    - `src/exstruct/core/pipeline.py`
    - `src/exstruct/models/__init__.py`
  - tests:
    - `tests/core/test_libreoffice_backend.py`
    - `tests/core/test_libreoffice_smoke.py`
