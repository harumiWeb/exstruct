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

## 2026-03-20 issue #107 extraction CLI startup optimization

### Planning

- [x] Confirm issue `#107` details with `gh issue view 107`.
- [x] Read current extraction CLI code, related docs/specs, and relevant tests.
- [x] Classify ADR need for the public CLI contract change.
- [x] Add the issue `#107` working spec to `tasks/feature_spec.md`.
- [x] Add or update the ADR for extraction CLI runtime capability validation.
- [x] Refactor extraction CLI parser construction so it never probes COM availability.
- [x] Always register `--auto-page-breaks-dir` in extraction CLI help.
- [x] Add runtime validation for `--auto-page-breaks-dir` covering `light` mode and unavailable COM on `standard` / `verbose`.
- [x] Keep existing `libreoffice` rejection and combined-error precedence intact.
- [x] Update extraction CLI tests for no-probe help and runtime validation.
- [x] Update current user docs and internal specs for the new CLI contract.
- [x] Run targeted pytest for extraction CLI coverage.
- [x] Run `uv run task build-docs`.
- [x] Run `uv run task precommit-run`.
- [x] Review task/spec retention and record permanent destinations in the Review section.

### Review

- Extraction CLI parser construction is now side-effect free: `build_parser()` always registers `--auto-page-breaks-dir`, and COM probing happens only when that flag is requested at runtime from a supported mode.
- `mode="light"` now fails explicitly for `--auto-page-breaks-dir`, while `mode="libreoffice"` keeps the existing core validation and combined-error precedence.
- Permanent destinations:
  - `dev-docs/adr/ADR-0008-extraction-cli-runtime-capability-validation.md` records the policy change and why parser-time probing is forbidden.
  - `docs/cli.md`, `README.md`, and `README.ja.md` now describe the always-visible flag and execution-time validation contract.
  - `dev-docs/specs/excel-extraction.md` now records the internal guarantee that the extraction CLI validates auto page-break export at runtime instead of parser construction.
- ADR checks:
  - `adr-linter`: no high/medium/low findings for `ADR-0008`.
  - `adr-reviewer`: `ready`, no findings.
  - `adr-reconciler`: no policy drift, no stale references, two low-evidence findings addressed by adding `verbose` and `main(["--help"])` regression tests.
  - `adr-indexer`: index artifacts were synchronized manually from the ADR source text (`README.md`, `index.yaml`, `decision-map.md`).
- Verification:
  - `gh issue view 107 --json number,title,body,labels,assignees,state,url`
  - `uv run pytest tests/cli/test_cli.py tests/cli/test_edit_cli.py -q`
  - `uv run task build-docs`
  - `uv run task precommit-run`
  - `git diff --check`

## 2026-03-20 issue #107 review follow-up: libreoffice auto page-break fast-fail

### Planning

- [x] Re-read the CLI review comment and compare it with `src/exstruct/cli/main.py` and the shared LibreOffice validator.
- [x] Update `tasks/feature_spec.md` with the follow-up contract and verification scope.
- [x] Make the CLI reject `--mode libreoffice --auto-page-breaks-dir` before `process_excel()`.
- [x] Preserve the existing combined LibreOffice error precedence when rendering flags are also present.
- [x] Add regression tests that prove the CLI rejects these requests before `process_excel()` runs.
- [x] Run targeted pytest for `tests/cli/test_cli.py`.
- [x] Update this Review section with the final validation result and retention decision.

### Review

- The review finding was valid: `_validate_auto_page_breaks_request()` treated `mode="light"` as a CLI-side fast-fail but let `mode="libreoffice"` fall through to the engine layer, which made responsibility and failure timing inconsistent.
- `src/exstruct/cli/main.py` now reuses `validate_libreoffice_process_request(...)` for `--auto-page-breaks-dir` in `mode="libreoffice"`, so the CLI rejects invalid requests before `process_excel()` while preserving the existing single-error and combined-error message precedence.
- `tests/cli/test_cli.py` now proves both `libreoffice + --auto-page-breaks-dir` and `libreoffice + --pdf + --auto-page-breaks-dir` fail before `process_excel()` runs, and that these paths do not probe COM availability.
- Retention decision:
  - No new permanent document was needed. This follow-up only brought the implementation back into alignment with the policy already recorded in `ADR-0008`, `docs/cli.md`, `README.md`, `README.ja.md`, and `dev-docs/specs/excel-extraction.md`.
  - The temporary working notes for this follow-up can stay limited to this section in `tasks/feature_spec.md` and `tasks/todo.md`.
- Verification:
  - `uv run pytest tests/cli/test_cli.py -q`
  - `uv run task precommit-run`
  - `git diff --check`

## 2026-03-20 issue #107 review follow-up: wording and help-text clarity

### Planning

- [x] Retrieve the PR review comments with `gh` and classify which findings are substantive.
- [x] Confirm the wording nits in `tasks/todo.md` and `dev-docs/adr/ADR-0008-extraction-cli-runtime-capability-validation.md`.
- [x] Clarify the `--auto-page-breaks-dir` help text in `src/exstruct/cli/main.py` so it matches the runtime contract.
- [x] Update CLI help tests for the clarified wording.
- [x] Run targeted pytest for `tests/cli/test_cli.py`.
- [x] Run `uv run task precommit-run`.
- [x] Record the review outcome and retention decision here.

### Review

- The explicit PR review findings were valid but minor: `tasks/todo.md` and `dev-docs/adr/ADR-0008-extraction-cli-runtime-capability-validation.md` each had a hyphenation nit (`low-evidence`, `side-effect-free`).
- A separate suppressed Copilot note about `--auto-page-breaks-dir` help text was also substantively valid: the old string mentioned only LibreOffice rejection and omitted that output files follow `--format`, while the actual runtime contract also rejects `light` and requires `standard`/`verbose` with Excel COM.
- `src/exstruct/cli/main.py` now states the fuller contract in the argument help text, and `tests/cli/test_cli.py` now checks for the clarified help wording without depending on exact `argparse` line wrapping.
- Retention decision:
  - No new ADR or spec migration was needed. The durable contract remains in `ADR-0008`, `docs/cli.md`, and the README files; this follow-up only aligns wording and help text with that existing policy.
  - The temporary working record can stay limited to this section in `tasks/feature_spec.md` and `tasks/todo.md`.
- Verification:
  - `gh pr view 111 --json number,title,reviewDecision,reviews,comments,files,url`
  - `gh api repos/harumiWeb/exstruct/pulls/111/comments`
  - `uv run pytest tests/cli/test_cli.py -q`
  - `uv run task precommit-run`
  - `git diff --check`

## 2026-03-20 issue #108 CLI startup lazy import optimization

### Planning

- [x] Confirm issue `#108` details with `gh issue view 108`.
- [x] Inspect current import paths in `src/exstruct/__init__.py`, `src/exstruct/edit/__init__.py`, `src/exstruct/cli/main.py`, and `src/exstruct/cli/edit.py`.
- [x] Classify ADR need for the startup optimization work.
- [x] Add the issue `#108` working spec to `tasks/feature_spec.md`.
- [x] Refactor `src/exstruct/__init__.py` to defer heavy imports while preserving exported symbol names.
- [x] Refactor `src/exstruct/edit/__init__.py` to defer heavy imports while preserving exported symbol names.
- [x] Refactor `src/exstruct/cli/main.py` so edit/extraction implementations load only after routing is known.
- [x] Refactor `src/exstruct/cli/edit.py` so `ops` commands avoid extraction-path imports and handler-specific dependencies load lazily.
- [x] Add regression tests for startup import isolation and existing CLI behavior.
- [x] Update `dev-docs/architecture/overview.md` with the lightweight-startup import rule.
- [x] Run targeted pytest for CLI/startup coverage.
- [x] Run `uv run task precommit-run`.
- [x] Record final verification and retention notes in this Review section.

### Review

- `src/exstruct/__init__.py` now keeps the public export surface but resolves heavy extraction/runtime symbols lazily, so importing `exstruct` no longer front-loads extraction engine modules for CLI startup.
- `src/exstruct/edit/__init__.py` now resolves editing exports lazily, which lets CLI code import edit submodules without paying the full patch-service import cost up front.
- `src/exstruct/cli/main.py` now keeps monkeypatch-compatible wrappers for `process_excel`, `get_com_availability`, `is_edit_subcommand`, and `run_edit_cli`, but loads the underlying implementations only after routing demands them.
- `src/exstruct/cli/edit.py` now keeps monkeypatch-compatible wrappers for `patch_workbook`, `make_workbook`, `resolve_top_level_sheet_for_payload`, and `validate_input`, while `ops` commands load only schema metadata and avoid dragging the extraction path into startup.
- `tests/cli/test_cli_lazy_imports.py` now locks the startup boundary with subprocess `sys.modules` probes for `import exstruct`, `import exstruct.cli.main`, `import exstruct.cli.edit`, and `main(["ops", "list"])`.
- `dev-docs/architecture/overview.md` now records the durable rule that package `__init__` files and lightweight CLI startup paths must remain side-effect-free.
- Retention decision:
  - No new ADR was added. The change preserves the public contract and only changes import timing, so the durable guidance lives in `dev-docs/architecture/overview.md`.
  - The temporary working record for implementation order and verification remains limited to this section in `tasks/feature_spec.md` and `tasks/todo.md`.
- Verification:
  - `uv run pytest tests/cli/test_cli.py tests/cli/test_edit_cli.py tests/cli/test_cli_lazy_imports.py tests/edit/test_architecture.py -q`
  - `uv run task precommit-run`
  - manual `-X importtime` sanity probe for `-m exstruct.cli.main --help` and `-m exstruct.cli.main ops list`

## 2026-03-20 issue #108 review and Codacy follow-up

### Planning

- [x] Retrieve PR `#112` Codacy findings and review comments with `scripts/codacy_issues.py` and `gh`.
- [x] Classify which findings are substantive and confirm the current implementation gaps locally.
- [x] Add the follow-up spec and task record to `tasks/feature_spec.md` and `tasks/todo.md`.
- [x] Replace the generic lazy-import helpers in `src/exstruct/__init__.py`, `src/exstruct/edit/__init__.py`, and `src/exstruct/cli/edit.py` with explicit literal loaders.
- [x] Restore runtime-resolvable type hints for public helpers in `src/exstruct/__init__.py` without eagerly importing `exstruct.models`.
- [x] Add a fast path in `src/exstruct/cli/main.py` so non-edit argv does not import `exstruct.cli.edit`.
- [x] Remove the top-level `pydantic` import from `src/exstruct/cli/edit.py`.
- [x] Add or update regression tests for startup import boundaries and runtime type hints.
- [x] Run targeted pytest for CLI follow-up coverage.
- [x] Run `uv run task precommit-run`.
- [x] Update this Review section with the final verification result and retention decision.

### Review

- Codacy's three `non-literal-import` findings were not exploitable security bugs in practice, because the module targets were fixed by code rather than user input. Even so, the finding was operationally valid for CI, so the generic loaders were replaced with explicit literal loader functions in `src/exstruct/__init__.py`, `src/exstruct/edit/__init__.py`, and `src/exstruct/cli/edit.py`.
- The PR review about runtime type hints was valid. `typing.get_type_hints(exstruct.extract)` regressed with `NameError` after the lazy-import refactor, so `src/exstruct/__init__.py` now patches the affected public helper annotations to resolve exported model types through `_lazy_type(...)` only when runtime introspection asks for them.
- The PR review about `cli.main` routing was valid. `src/exstruct/cli/main.py` now fast-fails obvious non-edit argv before importing `exstruct.cli.edit`, so `exstruct --help` and extraction-style argv no longer pay the edit-module import cost.
- The PR review about `pydantic` eager import in `src/exstruct/cli/edit.py` was valid for routing/help-only paths. The module now defers `pydantic` loading until an actual validation-error check happens and serializes JSON payloads via `model_dump` duck typing.
- `tests/cli/test_cli_lazy_imports.py` now locks the new boundaries: `import exstruct.cli.edit` keeps `pydantic` unloaded, `main(["--help"])` keeps `exstruct.cli.edit` unloaded, and `typing.get_type_hints(exstruct.extract)` resolves `WorkbookData` successfully.
- Retention decision:
  - No new ADR or permanent spec migration was needed. This follow-up only hardens the existing issue `#108` implementation and review expectations under the already-recorded lightweight-startup rule in `dev-docs/architecture/overview.md`.
  - The temporary working notes for this follow-up can remain limited to this section in `tasks/feature_spec.md` and `tasks/todo.md`.
- Verification:
  - `python scripts/codacy_issues.py --pr 112 --min-level Error`
  - `gh pr view 112 --json number,title,reviewDecision,reviews,comments,files,url,headRefName,baseRefName`
  - `gh api repos/harumiWeb/exstruct/pulls/112/comments`
  - `uv run pytest tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/cli/test_cli.py -q`
  - `uv run task precommit-run`
  - manual `uv run python` probes for `typing.get_type_hints(exstruct.extract)` and `main(["--help"])` import boundaries

## 2026-03-21 issue #108 review follow-up: validate runtime error scope

### Planning

- [x] Retrieve the new PR `#112` review comments and classify which ones are substantively valid.
- [x] Confirm locally whether `isinstance(exc, OSError | RuntimeError | ValueError)` is actually invalid on the supported Python runtime.
- [x] Add the working spec and task record for this follow-up.
- [x] Narrow `validate` exception handling in `src/exstruct/cli/edit.py` back to the original `(OSError, ValidationError, ValueError)` scope.
- [x] Add a regression test that proves `validate` still propagates `RuntimeError`.
- [x] Run targeted pytest for `tests/cli/test_edit_cli.py`.
- [x] Run `uv run task precommit-run`.
- [x] Update this Review section with the final verification result and retention decision.

### Review

- The new Devin review finding was valid: the shared `_is_cli_runtime_error(...)` helper widened `_run_validate_command(...)` to catch `RuntimeError`, which changed the historical validate-subcommand contract.
- The new Copilot review finding was not valid on the supported runtime. A direct `uv run python` probe confirmed that `isinstance(OSError(), OSError | RuntimeError | ValueError)` evaluates successfully on Python `3.11`, so no change was made for that comment.
- `src/exstruct/cli/edit.py` now uses a separate `_is_validate_cli_error(...)` helper so `patch` / `make` still catch `RuntimeError` while `validate` only catches `(OSError, ValidationError, ValueError)` as before.
- `tests/cli/test_edit_cli.py` now includes a regression test proving that `validate` propagates `RuntimeError` instead of converting it to handled CLI stderr output.
- Retention decision:
  - No new ADR or permanent spec migration was needed. This follow-up only restores the pre-existing validate CLI error boundary inside the current edit CLI design.
  - The temporary working notes for this review follow-up can remain limited to this section in `tasks/feature_spec.md` and `tasks/todo.md`.
- Verification:
  - `gh api repos/harumiWeb/exstruct/pulls/112/comments`
  - `gh api graphql -f query='query { repository(owner:"harumiWeb", name:"exstruct") { pullRequest(number: 112) { reviewThreads(first: 30) { nodes { id isResolved isOutdated comments(first: 20) { nodes { id author { login } body path url createdAt } } } } } } }'`
  - `uv run python` probe for `isinstance(OSError(), OSError | RuntimeError | ValueError)`
  - `uv run pytest tests/cli/test_edit_cli.py -q`
  - `uv run task precommit-run`
