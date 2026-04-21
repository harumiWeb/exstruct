# Todo

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
- [ ] Build the pure-Python rich extraction baseline from OOXML parsing.
- [ ] Improve pure-Python geometry fidelity for shapes/connectors/charts on sheets with custom row heights or column widths.
- [x] Decide that the new capability is exposed by strengthening `light` itself.
- [ ] Add regression coverage for OOXML-only rich extraction and optional LibreOffice enrichment.
- [ ] Update ADR/spec/docs/schema artifacts once the policy decision is accepted.

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
