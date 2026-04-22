# Todo

## 2026-04-22 README English/Japanese parity refresh

### Planning

- [x] Compare `README.ja.md` against `README.md` and identify English-only sections/details that should be removed.
- [x] Update `README.md` so its structure and examples match the edited Japanese README.
- [x] Run a lightweight documentation verification pass and record the result.

### Review

- `README.md` now follows the same top-level structure as `README.ja.md`, including the language switcher, quick-start ordering, sample sections, and closing reference sections.
- Removed English-only content that the Japanese README no longer keeps, including the `Choose an Interface` section, extra MCP operational notes, and extra positioning commentary around editing workflows.
- Reworked the English README intro, feature list, installation notes, CLI/MCP guidance, and example assets so they match the edited Japanese README more closely.
- Expanded the English version of Example 2 so the LLM-output section now covers the same spouse / income / asset / applicant subsections that remain in `README.ja.md`.
- Verification:
  - `rg -n '^#{1,6} ' README.md README.ja.md`
  - `git diff --check -- README.md tasks/feature_spec.md tasks/todo.md`
  - result: passed

## 2026-04-22 light-mode print areas / OOXML drawing resilience

### Planning

- [x] Align `process_excel` / engine auto-filter defaults with the accepted `light` print-area contract.
- [x] Make `read_sheet_drawings()` skip only malformed sheets instead of dropping the whole workbook.
- [x] Add regression coverage for light-mode print-area side outputs and partial OOXML drawing failure.
- [x] Regenerate generated model docs and run verification.

### Review

- `process_excel()`, CLI, and engine export now keep `print_areas` in `light` mode by default, matching the accepted `ADR-0010` and current public docs.
- `read_sheet_drawings()` now degrades per sheet: malformed drawing XML on one worksheet logs a warning and skips that sheet without erasing healthy OOXML-rich artifacts from sibling sheets.
- Added regression coverage for:
  - `light` engine export writing `print_areas_dir`
  - `process_excel(..., mode="light", print_areas_dir=...)`
  - `exstruct --mode light --print-areas-dir ...`
  - direct OOXML parser behavior when one sheet drawing is malformed
- Permanent-document follow-up:
  - updated `dev-docs/testing/test-requirements.md` `MODE-08`
  - regenerated `docs/generated/models.md` for the `FilterOptions.include_print_areas` auto-description
- Verification:
  - `uv run pytest tests/engine/test_engine.py tests/core/test_mode_output.py tests/cli/test_cli.py tests/core/test_ooxml_drawing.py -q`
    - result: `49 passed, 3 skipped`
  - `uv run python scripts/gen_model_docs.py`
    - result: passed
  - `uv run task precommit-run`
    - result: passed

## 2026-04-22 PR #129 review follow-up

### Planning

- [x] Inspect unresolved PR review threads and cluster the actionable comments by code path.
- [x] Fix the `process_excel()` filter default, light-pipeline fallback handling, and stale pipeline docs.
- [x] Reduce OOXML sheet-metrics overhead with cached offsets and a streaming worksheet metrics reader.
- [x] Add or update regression tests for the review follow-ups.
- [x] Run targeted pytest coverage and `uv run task precommit-run`.

### Review

- Addressed all six unresolved actionable review threads on PR `#129`.
- `process_excel()` now leaves `FilterOptions.include_print_areas=None`, so it follows the engine's auto-default contract instead of re-hardcoding print-area inclusion.
- `_run_light_pipeline()` now degrades through the existing fallback path when OOXML rich extraction raises unexpectedly, while preserving already extracted rich artifacts when only the chart step fails.
- `SheetDrawingMetrics` now caches cumulative row/column offsets, and `_read_sheet_metrics()` now streams worksheet XML with `iterparse()` so drawing geometry no longer requires a full worksheet DOM parse.
- `dev-docs/architecture/pipeline.md` now documents `OoxmlRichBackend` as the concrete `RichBackend` used by `light` mode.
- Added regression coverage for:
  - `process_excel()` keeping the engine auto print-area default
  - light-mode fallback when chart extraction fails
  - out-of-order cached row/column offset lookups
- Verification:
  - `uv run pytest tests/core/test_pipeline.py tests/core/test_mode_output.py tests/core/test_ooxml_drawing.py -q`
    - result: `69 passed`
    - note: pytest still emitted the pre-existing Windows COM fatal-exception noise after success, but exited with code `0`
  - `uv run task precommit-run`
    - result: passed

## 2026-04-22 PR #129 review follow-up (second pass)

### Planning

- [x] Re-check unresolved review threads after commit `14efdda` and separate already-fixed/outdated items from new actionable findings.
- [x] Fix the remaining per-sheet OOXML drawing exception hole and stale `README` wording about `light` / non-COM extraction.
- [x] Normalize the specific YAML/Markdown files flagged for CRLF line endings back to LF.
- [x] Run the relevant pytest coverage and `uv run task precommit-run`.
- [x] Push the follow-up commit to refresh PR `#129`.

### Review

- Re-checked the unresolved PR threads after commit `14efdda` and confirmed the earlier `process_excel`, light-pipeline fallback, metrics caching, and architecture-doc comments were already addressed; the remaining actionable items were one `BadZipFile` exception gap, stale README wording, and newline-only cleanup.
- `read_sheet_drawings()` now treats `BadZipFile` the same as the other sheet-local drawing parse failures, so one corrupt drawing member no longer forces workbook-wide OOXML rich artifacts to `{}` through the outer backend fallback.
- `README.md` and `README.ja.md` now describe the current `light` contract consistently in the CLI quick start, non-COM note, output-mode section, and fallback section.
- Normalized the specific files flagged by review bots back to LF:
  - `.agents/skills/exstruct-cli/agents/openai.yaml`
  - `dev-docs/agents/coding-guidelines.md`
  - `mkdocs.yml`
- Added regression coverage for the per-sheet `BadZipFile` case in `tests/core/test_ooxml_drawing.py`.
- Verification:
  - `uv run pytest tests/core/test_ooxml_drawing.py -q`
    - result: `7 passed`
  - `uv run task precommit-run`
    - result: passed


## 2026-03-19 v0.7.0 release closeout

### Planning

- [x] Add the `0.7.0` changelog entry with `Added` / `Changed` / `Fixed`.
- [x] Create `docs/release-notes/v0.7.0.md` for issue `#99` and maintainer-facing docs changes.
- [x] Add `v0.7.0` to the `Release Notes` nav in `mkdocs.yml`.
- [x] Align the local package version in `pyproject.toml` and `uv.lock` to `0.7.0`.
- [x] Move the legacy monkeypatch compatibility precedence note to `dev-docs/architecture/overview.md`.
- [x] Compress `tasks/feature_spec.md` and `tasks/todo.md` to a single release-closeout section.
- [x] Run `uv run task build-docs`.
- [x] Run `uv run task precommit-run`.
- [x] Run the release-prep `rg` and `git diff --check` consistency checks.

### Review

- `CHANGELOG.md`, `docs/release-notes/v0.7.0.md`, `mkdocs.yml`, `pyproject.toml`, and `uv.lock` now describe and label the `0.7.0` release consistently.
- `dev-docs/architecture/overview.md` now holds the maintainer-facing note that compatibility shims must preserve live monkeypatch visibility and verify override precedence at the highest public entrypoint.
- Historical issue `#99`, PR follow-up, and prior cleanup logs were intentionally removed from `tasks/feature_spec.md` and `tasks/todo.md` after permanent information was classified and retained in `CHANGELOG.md`, `docs/`, and `dev-docs/architecture/`.
- No new `dev-docs/specs/` or `dev-docs/adr/` migration was required for this closeout; `ADR-0006`, `ADR-0007`, and the editing specs remain the canonical permanent sources.
- Verification:
  - `uv run task build-docs`
  - `uv run task precommit-run`
  - `rg -n "0\.7\.0|v0\.7\.0" pyproject.toml uv.lock CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.0.md`
  - `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
  - `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.0.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md dev-docs/architecture/overview.md`

## 2026-03-21 v0.7.1 release closeout

### Planning

- [x] Add the `0.7.1` changelog entry with `Added` / `Changed` / `Fixed`.
- [x] Create `docs/release-notes/v0.7.1.md` for issues `#107`, `#108`, and `#109`.
- [x] Add `v0.7.1` to the `Release Notes` nav in `mkdocs.yml`.
- [x] Align the local package version in `pyproject.toml` and the editable `exstruct` package entry in `uv.lock` to `0.7.1`.
- [x] Compress the detailed `#107` / `#108` working logs in `tasks/feature_spec.md` and `tasks/todo.md` into this release-closeout record.
- [x] Run `uv run pytest tests/cli/test_cli.py tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`.
- [x] Run `uv run task build-docs`.
- [x] Run `uv run task precommit-run`.
- [x] Run the release-prep `rg` and `git diff --check` consistency checks.

### Review

- `CHANGELOG.md`, `docs/release-notes/v0.7.1.md`, `mkdocs.yml`, `pyproject.toml`, and `uv.lock` now describe and label the `0.7.1` release consistently around the CLI/package import optimization work from issues `#107`, `#108`, and `#109`.
- The release narrative explicitly documents the public behavior deltas that shipped after `v0.7.0`: runtime validation for `--auto-page-breaks-dir`, lighter startup/import behavior for CLI and package entrypoints, preserved exported symbol names, and the restored `validate` error boundary.
- Historical implementation and review logs for issues `#107` and `#108` were intentionally removed from `tasks/feature_spec.md` and `tasks/todo.md` after permanent information was classified and retained in `CHANGELOG.md`, `docs/release-notes/v0.7.1.md`, `docs/cli.md`, `README.md`, `README.ja.md`, `dev-docs/specs/excel-extraction.md`, `dev-docs/architecture/overview.md`, and `ADR-0008`.
- No new `dev-docs/specs/` or `dev-docs/adr/` migration was required for this closeout; the existing CLI docs, architecture note, extraction spec, and `ADR-0008` remain the canonical permanent sources for the shipped behavior.
- Verification:
  - `uv run pytest tests/cli/test_cli.py tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`
  - `uv run task build-docs`
  - `uv run task precommit-run`
  - `rg -n "0\.7\.1|v0\.7\.1" CHANGELOG.md mkdocs.yml docs/release-notes/v0.7.1.md`
  - `rg -n '^version = "0\.7\.1"$' pyproject.toml uv.lock`
  - `rg -n "^## " tasks/feature_spec.md tasks/todo.md`
  - `git diff --check -- CHANGELOG.md docs/release-notes/v0.7.1.md mkdocs.yml pyproject.toml uv.lock tasks/feature_spec.md tasks/todo.md`

## 2026-03-21 issue #115 ExStruct CLI Skill

### Planning

- [x] Confirm the ADR verdict and draft the new policy record for the CLI Skill packaging and CLI/MCP boundary decisions.
- [x] Add an internal spec under `dev-docs/specs/` for the `exstruct-cli` Skill contract.
- [x] Create `.agents/skills/exstruct-cli/` with `SKILL.md`, `agents/openai.yaml`, and the required `references/` files.
- [x] Keep `SKILL.md` concise and move detailed command, verification, and backend guidance into `references/`.
- [x] Update `README.md` and `README.ja.md` with Skill installation guidance, when-to-use guidance, and minimal example prompts.
- [x] Run the external `skill-creator` helper scripts to generate and validate `agents/openai.yaml`.
- [x] Run repo verification (`uv run task precommit-run`) and record the result.
- [x] Review which #115 notes remain temporary versus durable, then compress the `tasks/` sections after the durable content is migrated.

### Review

- Introduced the canonical repo-owned Skill at `.agents/skills/exstruct-cli/` with a lean `SKILL.md`, five focused reference documents, and generated `agents/openai.yaml`.
- Published `dev-docs/specs/exstruct-cli-skill.md` as the durable contract for the Skill layout, trigger/positioning rules, and verification obligations.
- Documented the policy decision in `dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md` and synchronized `dev-docs/adr/README.md`, `dev-docs/adr/index.yaml`, and `dev-docs/adr/decision-map.md`.
- Expanded `README.md` and `README.ja.md` with installation guidance, CLI-vs-MCP usage boundaries, and example prompts.
- The durable content for #115 now lives in `.agents/skills/exstruct-cli/`, `dev-docs/specs/exstruct-cli-skill.md`, `dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md`, `README.md`, and `README.ja.md`; this `tasks/` section remains only as a compact implementation record.
- Verification:
  - `python <skill-creator>/scripts/generate_openai_yaml.py .agents/skills/exstruct-cli --interface 'display_name=ExStruct CLI' --interface 'short_description=Guide safe ExStruct CLI edit workflows' --interface 'default_prompt=Use $exstruct-cli to choose the right ExStruct editing CLI command, follow a safe validate/dry-run workflow that inspects the PatchResult/diff before applying changes, and explain any backend constraints for this workbook task.'`
  - `python <skill-creator>/scripts/quick_validate.py .agents/skills/exstruct-cli`
  - `rg -n "ExStruct CLI Skill|exstruct-cli|validate -> dry-run -> inspect -> apply -> verify|\.agents/skills/exstruct-cli" README.md README.ja.md dev-docs/specs/exstruct-cli-skill.md dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md .agents/skills/exstruct-cli/ -g "*"`
  - `rg -n "^## |Tests:|Code:|Related specs:" dev-docs/adr/ADR-0009-single-cli-skill-for-agent-workflows.md`
  - `uv run task precommit-run`
  - `git diff --check`

## 2026-04-16 SECURITY.md policy

### Planning

- [x] Confirm whether `SECURITY.md` already exists and review the current public-document tone in `README.md` and `CONTRIBUTING.md`.
- [x] Define the minimal public policy: latest-release-only support and email-first disclosure to `harumiweb.security@gmail.com`.
- [x] Add a root-level `SECURITY.md` with supported versions, reporting instructions, and expectations for response/disclosure.
- [x] Record the durable destination and ADR verdict in `tasks/feature_spec.md`.
- [x] Run the planned verification commands and record the results.

### Review

- Added the root-level `SECURITY.md` as the durable public security policy document with latest-release-only support guidance and email-first reporting to `harumiweb.security@gmail.com`.
- Kept the change documentation-only; `README.md`, `README.ja.md`, `docs/`, `mkdocs.yml`, code, and public runtime interfaces were unchanged.
- `tasks/feature_spec.md` now records the compact spec, permanent destination, verification commands, and `not-needed` ADR verdict for this session.
- Verification:
  - `rg -n "Security Policy|harumiweb.security@gmail.com|Latest release|GitHub Issues" SECURITY.md`
  - `git diff --check -- SECURITY.md tasks/feature_spec.md tasks/todo.md`
  - `uv run task precommit-run`
  - `uv run pytest -q`
  - Result summary: `pre-commit` passed (`ruff`, `ruff-format`, `mypy`), and `pytest` completed with `913 passed, 4 skipped`.

## 2026-04-16 issue #77 LibreOffice typed workbook handle

### Planning

- [x] Confirm the issue scope, related PR review comment, and ADR verdict for the typed handle change.
- [x] Record the compact spec and implementation constraints in `tasks/feature_spec.md`.
- [x] Replace the raw workbook token in `src/exstruct/core/libreoffice.py` with a typed handle and meaningful `close_workbook()` cleanup.
- [x] Update `src/exstruct/core/backends/libreoffice_backend.py` and test doubles to use the typed workbook lifecycle.
- [x] Add regression tests for typed handle ownership, idempotent close, and cache invalidation.
- [x] Run the targeted LibreOffice backend tests and `uv run task precommit-run`.
- [x] Fill the Review section with verification evidence and permanent-destination notes.

### Review

- Added `LibreOfficeWorkbookHandle` in `src/exstruct/core/libreoffice.py` and replaced the raw `dict` workbook token with a typed, session-owned handle.
- `LibreOfficeSession.close_workbook()` now validates session ownership, is idempotent for repeated close calls, and clears session-local bridge cache entries when the last handle for a workbook is released.
- `LibreOfficeSession.extract_draw_page_shapes()` and `extract_chart_geometries()` now accept either a direct `Path` or a typed workbook handle, preserving the existing path-based call pattern while enabling the typed lifecycle.
- `src/exstruct/core/backends/libreoffice_backend.py` now prefers `load_workbook() -> extract_*() -> close_workbook()` when the session implements that lifecycle, while preserving the legacy path-only `session_factory` extension point as a runtime fallback.
- `tests/core/test_libreoffice_backend.py` now covers typed handle creation, backend lifecycle usage, legacy path-only session compatibility, foreign-session rejection, repeated close idempotence, closed-handle extraction failure, and cache invalidation after close.
- Review follow-up: `session_factory` is now typed as a structural path-or-lifecycle protocol instead of the concrete built-in session, so legacy custom sessions and test doubles no longer need `cast(LibreOfficeSession, ...)` to type-check.
- Review follow-up: workbook-handle validation now rejects rehydrated handles whose `file_path` disagrees with the registered workbook id, preventing forged handles from reusing another workbook's cache/close path.
- Review follow-up: `_resolve_workbook_path()` now directly returns `_require_handle_path()` for typed handles, removing an unreachable `None` branch so the control flow matches the actual closed-handle behavior.
- No new `dev-docs/specs/`, `dev-docs/architecture/`, `dev-docs/adr/`, or public `docs/` updates were required; this issue only hardened the existing internal LibreOffice session contract.
- Verification:
  - `uv run pytest tests/core/test_libreoffice_backend.py -q`
  - `uv run task precommit-run`

## 2026-04-21 Pure-Python rich extraction for light-mode environments

### Planning

- [x] Review the current non-COM extraction contract in `docs/`, `dev-docs/specs/`, ADRs, and task guidance before proposing backend changes.
- [x] Inspect the current LibreOffice-mode code path in `src/exstruct/core/backends/libreoffice_backend.py`, `src/exstruct/core/ooxml_drawing.py`, `src/exstruct/core/libreoffice.py`, and the related tests.
- [x] Identify which rich artifacts already come from pure OOXML and which ones still depend on LibreOffice runtime enrichment.
- [x] Re-scope the task after the user clarified that the goal is richer extraction in light-mode environments, not removing LibreOffice dependency from the existing LibreOffice path.
- [x] Decide the ADR verdict and likely permanent-document destinations for the eventual implementation.
- [x] Draft `ADR-0010` for changing `light` into the pure-Python OOXML-rich baseline and link the ADR index artifacts.
- [x] Run an ADR-linter-style structural check on the draft and confirm the supersede/index links are consistent.
- [x] Build the pure-Python rich extraction baseline from OOXML parsing.
- [x] Improve pure-Python geometry fidelity for shapes/connectors/charts on sheets with custom row heights or column widths.
- [x] Decide that the new capability is exposed by strengthening `light` itself.
- [x] Add regression coverage for OOXML-only rich extraction and optional LibreOffice enrichment.
- [x] Update ADR/spec/docs/public contract artifacts once the policy decision is accepted.

### Review

- Investigation result: the repository already has most of the needed Python-only parsing logic.
  - `src/exstruct/core/ooxml_drawing.py` already parses shapes, connectors, and charts from OOXML drawing parts.
  - `src/exstruct/core/backends/libreoffice_backend.py` already has `_build_shapes_from_ooxml(...)`, and chart extraction already builds metadata from OOXML before optionally refining geometry with LibreOffice.
- Root cause: the pure-Python capability exists in pieces, but the current non-COM product path does not expose it to environments that only receive `light`-level extraction.
- Main implementation risk: OOXML geometry is not yet strong enough to replace LibreOffice geometry on its own because `_marker_to_points()` still relies on fixed default row/column sizes and does not model custom sheet dimensions.
- Recommended rollout:
  - first build the pure-Python rich baseline
  - then tighten geometry fidelity and regression coverage
  - then wire that baseline into `light` and keep `libreoffice` as the optional enrichment layer
  - finally update ADR/spec/docs/schema artifacts once the contract details, especially mode exposure and `provenance`, are settled
- ADR verdict for the future code change: `required`.
  - related ADRs: `ADR-0001` and `ADR-0002`
  - draft created: `dev-docs/adr/ADR-0010-light-mode-as-the-pure-python-rich-ooxml-baseline.md`
- ADR draft check:
  - no unresolved structural holes were found in `ADR-0010`; required sections and evidence headings are present
  - supersede references are linked in `ADR-0010`, `ADR-0001`, `dev-docs/adr/README.md`, `dev-docs/adr/index.yaml`, and `dev-docs/adr/decision-map.md`
  - residual risk: `ADR-0001` still has status `accepted` while `ADR-0010` is only `proposed`; the final status flip should happen when `ADR-0010` is accepted
- Permanent-document note:
  - this section is still a temporary planning record
  - the draft policy now lives in `ADR-0010`; once implementation starts, the durable follow-up must update `dev-docs/specs/excel-extraction.md` and, if metadata changes, `dev-docs/specs/data-model.md` / `schemas/` / public docs
- Implementation status after the first code pass:
  - added `src/exstruct/core/backends/ooxml_backend.py` as the pure-Python OOXML rich backend
  - `light` now runs that backend and keeps shapes/charts in the assembled workbook instead of forcing the old empty-rich-artifact fallback
  - `libreoffice` now seeds the same OOXML baseline first, so runtime-unavailable fallback preserves `python_ooxml` shapes/charts when they are available from OOXML
  - public model/schema provenance now includes `python_ooxml`
- Geometry-fidelity follow-up completed in the second code pass:
  - added worksheet-driven `SheetDrawingMetrics` parsing in `src/exstruct/core/ooxml_drawing.py` so anchor fallback uses sheet XML row heights and column widths instead of fixed defaults alone
  - shape and connector placement now prefers `a:xfrm` absolute position when the transform carries a non-zero size, while chart frames still fall back cleanly to anchor geometry when `xfrm` is a zero placeholder
  - added focused regression coverage in `tests/core/test_ooxml_drawing.py` for custom row/column metrics, two-cell-anchor geometry, and `sample/flowchart/sample-shape-connector.xlsx` transform placement
- Permanent-document follow-up completed:
  - `ADR-0010` is now `accepted`, `ADR-0001` is now `superseded`, and the ADR index artifacts (`README.md`, `index.yaml`, `decision-map.md`) were synchronized to match the source ADRs
  - `dev-docs/specs/excel-extraction.md`, `dev-docs/specs/data-model.md`, `dev-docs/testing/test-requirements.md`, `docs/cli.md`, `docs/api.md`, and `docs/mcp.md` now describe `light` as the pure-Python OOXML-rich baseline and document `python_ooxml` backend metadata
- Verification:
  - `uv run pytest tests/core/test_mode_output.py tests/core/test_pipeline.py tests/integration/test_integrate_raw_data.py tests/models/test_models_export.py tests/models/test_schemas_generated.py -q`
    - result: `75 passed, 2 skipped`
  - `uv run pytest tests/core/test_ooxml_drawing.py tests/core/test_mode_output.py tests/core/test_pipeline.py tests/integration/test_integrate_raw_data.py tests/models/test_models_export.py tests/models/test_schemas_generated.py -q`
    - result: `80 passed, 2 skipped`
  - `uv run pytest tests/core/test_ooxml_drawing.py tests/core/test_libreoffice_backend.py -q`
    - result: `59 passed`
  - `uv run task precommit-run`
    - result: passed after aligning `ComRichBackend` method signatures with the widened `RichBackend` protocol
  - `uv run task precommit-run`
    - result: passed after the geometry follow-up and durable docs updates
  - `uv run task build-docs`
    - result: still fails in `docs/api.md` mkdocstrings signature rendering with `AttributeError: 'NoneType' object has no attribute 'replace'`
    - baseline check: the same failure reproduces in a detached worktree at commit `c4d9acf`, so the docs-build failure is pre-existing and not introduced by this task
- Remaining work:
  - no task-local follow-up remains; the only unresolved item observed during verification is the pre-existing `build-docs` failure in mkdocstrings for `docs/api.md`
