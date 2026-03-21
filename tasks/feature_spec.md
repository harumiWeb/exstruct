# Feature Spec

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

## 2026-03-20 issue #107 extraction CLI startup optimization

### Goal

- Stop probing Excel COM availability while building the extraction CLI parser.
- Keep `--auto-page-breaks-dir` visible in help output on every host.
- Validate `--auto-page-breaks-dir` only when the user actually requests it at execution time.
- Return an explicit CLI error when auto page-break export is requested from an unsupported mode or unsupported runtime.

### Public contract

- `build_parser()` and `exstruct --help` must not call the COM availability probe.
- `--auto-page-breaks-dir` is always listed in extraction CLI help output.
- `--auto-page-breaks-dir` runtime behavior:
  - `mode="libreoffice"` keeps the existing `ConfigError` path and combined-error precedence.
  - `mode="light"` is rejected explicitly by the CLI with a message that auto page-break export requires `standard` or `verbose` with Excel COM.
  - `mode="standard"` / `mode="verbose"` trigger COM availability probing only when the flag is present.
  - When COM is unavailable, the CLI exits non-zero and prints an actionable message that names the flag and includes the availability reason when present.
- Existing `--pdf` / `--image` runtime behavior is unchanged in this task.

### Permanent destinations

- `dev-docs/adr/`
  - `ADR-0008` records the extraction CLI policy change: runtime capability validation instead of parser-time environment probing.
- `docs/`
  - `docs/cli.md` becomes the canonical public CLI contract for always-visible `--auto-page-breaks-dir` and runtime validation wording.
- `README.md` and `README.ja.md`
  - Update quick-start examples and current-behavior prose so they no longer claim the flag is hidden on unsupported hosts.
- `dev-docs/specs/`
  - `dev-docs/specs/excel-extraction.md` should keep the internal guarantee that auto page-break extraction is COM-only and tied to runtime validation.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Retain only the temporary working record, verification, and migration notes for this issue.

### Constraints

- Do not broaden the task into a Python API or MCP contract change.
- Do not rewrite historical release notes that describe the old behavior at the time they shipped.
- Preserve existing file-not-found behavior and existing `libreoffice` combined-error precedence.

### Verification plan

- `tests/cli/test_cli.py`
  - Help output always includes `--auto-page-breaks-dir`.
  - Parser/help generation does not call `get_com_availability()`.
  - `mode="light"` + `--auto-page-breaks-dir` fails with a clear runtime error.
  - `mode="standard"` / `mode="verbose"` + `--auto-page-breaks-dir` + unavailable COM fails with a clear runtime error.
  - Existing `libreoffice` rejection tests still pass.
- Targeted pytest for the CLI extraction test module.
- `uv run task build-docs`
- `uv run task precommit-run`

### ADR verdict

- `required`
- rationale: this changes the public extraction CLI contract for help visibility and execution-time validation of a COM-only flag, and it sets a reusable policy for future capability-gated CLI features.

## 2026-03-20 issue #107 review follow-up: libreoffice auto page-break fast-fail

### Goal

- Align `mode="libreoffice"` handling for `--auto-page-breaks-dir` with the new CLI-side fast-fail policy.
- Preserve the existing LibreOffice single-error and combined-error precedence from the shared validator.
- Prove that the CLI rejects the invalid request before `process_excel()` runs.

### Public contract

- `--mode libreoffice --auto-page-breaks-dir ...` fails in the CLI layer without calling `process_excel()`.
- When `--pdf` or `--image` is also present, the CLI keeps the existing combined LibreOffice error message precedence.
- This follow-up does not change the already documented `standard` / `verbose` COM-runtime validation policy.

### Permanent destinations

- No new permanent destination is required beyond the documents already updated for issue `#107`.
- The durable contract remains in `docs/cli.md`, `README.md`, `README.ja.md`, `dev-docs/specs/excel-extraction.md`, and `dev-docs/adr/ADR-0008-extraction-cli-runtime-capability-validation.md`.

### Constraints

- Reuse the existing LibreOffice validator instead of duplicating message composition in the CLI.
- Keep parser/help no-probe behavior unchanged.

### Verification plan

- `tests/cli/test_cli.py`
  - `mode="libreoffice"` + `--auto-page-breaks-dir` rejects before `process_excel()`.
  - `mode="libreoffice"` + rendering + `--auto-page-breaks-dir` keeps the combined error and also rejects before `process_excel()`.
- Targeted pytest for `tests/cli/test_cli.py`.

### ADR verdict

- `not-needed`
- rationale: this is a corrective follow-up that aligns implementation with the already-recorded `ADR-0008` policy rather than creating a new architectural decision.

## 2026-03-20 issue #107 review follow-up: wording and help-text clarity

### Goal

- Resolve the PR review wording nits in tracked documentation.
- Make the extraction CLI help text for `--auto-page-breaks-dir` match the runtime contract already documented elsewhere.

### Public contract

- The help text for `--auto-page-breaks-dir` states that it writes one file per auto page-break area, follows `--format`, and requires `--mode standard` or `--mode verbose` with Excel COM.
- This follow-up does not change runtime behavior; it only tightens wording and help-text clarity.

### Permanent destinations

- No new permanent destination is required.
- The durable wording lives in `src/exstruct/cli/main.py`, `dev-docs/adr/ADR-0008-extraction-cli-runtime-capability-validation.md`, and the existing issue `#107` task notes.

### Constraints

- Keep the change limited to wording/help-text clarity; do not change runtime validation or expand scope beyond the reviewed lines.

### Verification plan

- `tests/cli/test_cli.py`
  - Help output still includes `--auto-page-breaks-dir`.
  - Help output includes the clarified runtime-contract wording.
- `uv run pytest tests/cli/test_cli.py -q`

### ADR verdict

- `not-needed`
- rationale: this is a wording-only follow-up under the existing `ADR-0008` decision.

## 2026-03-20 issue #108 CLI startup lazy import optimization

### Goal

- Reduce startup import cost for lightweight CLI paths such as `exstruct --help` and `exstruct ops list`.
- Keep the existing extraction and editing CLI contracts unchanged while delaying heavy implementation imports until routing is known.
- Preserve the current module-level monkeypatch surfaces used by the CLI test suite.

### Public contract

- `exstruct --help` and parser construction keep the current CLI syntax, help text, and exit behavior.
- `exstruct ops list` and `exstruct ops describe` keep their current output shape and exit behavior.
- Extraction invocations still call `process_excel(...)` with the same arguments and keep current file-not-found and auto page-break validation behavior.
- Public Python symbol names exported from `exstruct` and `exstruct.edit` remain unchanged; only their import timing changes.

### Internal implementation guarantees

- `src/exstruct/__init__.py` must not eagerly import extraction engine, IO, render, or model modules during package import; convenience functions may import those dependencies inside function bodies.
- `src/exstruct/edit/__init__.py` must not eagerly import editing service/runtime/model modules during package import; exported names should resolve lazily.
- `src/exstruct/cli/main.py` must route edit vs extraction commands before importing edit/extraction implementations.
- `src/exstruct/cli/edit.py` must not import `exstruct.mcp.validate_input` or editing execution helpers at module import time; command handlers load only the functionality they need.
- Existing CLI module patch points (`process_excel`, `get_com_availability`, `is_edit_subcommand`, `run_edit_cli`, `patch_workbook`, `make_workbook`, `resolve_top_level_sheet_for_payload`, `validate_input`) remain present as thin wrappers.

### Scope and non-goals

- In scope:
  - `src/exstruct/__init__.py`
  - `src/exstruct/edit/__init__.py`
  - `src/exstruct/cli/main.py`
  - `src/exstruct/cli/edit.py`
  - targeted tests and one architecture note
- Out of scope:
  - changing CLI syntax, help wording, or JSON contracts
  - changing backend selection policy
  - optimizing `exstruct validate` startup beyond removing it from the `ops` path
  - refactoring `src/exstruct/mcp/__init__.py` unless required by failing tests on the `ops` path

### Permanent destinations

- `dev-docs/architecture/overview.md`
  - Records that package `__init__` files and lightweight CLI startup paths must remain side-effect-free and defer heavy imports.
- `tasks/feature_spec.md` and `tasks/todo.md`
  - Keep the temporary implementation/verification record for this issue.
- `dev-docs/adr/`
  - No new ADR is planned; this issue changes import timing only and does not alter the public contract or policy.

### Verification plan

- `tests/cli/test_cli.py`
  - help and extraction routing still behave the same
  - lightweight startup paths do not eagerly load edit/extraction implementation modules
- `tests/cli/test_edit_cli.py`
  - `ops list` / `ops describe` do not depend on extraction import paths
  - existing monkeypatch-based tests still pass
- `tests/edit/test_architecture.py` or a focused startup test module
  - `import exstruct` does not eagerly load extraction engine modules
  - `import exstruct.cli.edit` does not eagerly load `exstruct.mcp` / `exstruct.mcp.extract_runner`
- `uv run pytest tests/cli/test_cli.py tests/cli/test_edit_cli.py tests/edit/test_architecture.py -q`
- `uv run task precommit-run`
- manual importtime sanity checks for `--help` and `ops list`

### ADR verdict

- `not-needed`
- rationale: this is a startup-focused internal refactor that preserves existing CLI/API contracts and backend policy. The durable guidance belongs in architecture notes rather than a new policy ADR.

## 2026-03-20 issue #108 review and Codacy follow-up

### Goal

- Resolve the 3 Codacy `non-literal-import` findings on PR `#112` without regressing the lazy-import startup work.
- Address the substantive PR review comments about runtime annotation introspection and unnecessary eager imports on lightweight CLI paths.
- Keep the public CLI and Python export surface unchanged while tightening the internal implementation.

### Public contract

- `typing.get_type_hints(exstruct.extract)` and the other public convenience helpers in `src/exstruct/__init__.py` must keep resolving runtime-visible exported model types after the lazy-import refactor.
- `exstruct --help` and extraction-style argv that are clearly not edit subcommands must not import `exstruct.cli.edit`.
- Importing `exstruct.cli.edit` for routing/help-only purposes must not eagerly import `pydantic`.
- Public exports from `exstruct` and `exstruct.edit` remain unchanged; only the internal lazy-loader structure changes to satisfy static analysis.

### Constraints

- Do not undo the startup optimization by eagerly importing `exstruct.models`, `exstruct.edit.models`, or `pydantic` at module import time.
- Replace generic non-literal `import_module()` helpers with explicit literal import paths or literal loader functions so Codacy/Semgrep no longer flags them.
- Keep the existing monkeypatch-compatible wrappers in `src/exstruct/cli/main.py` and `src/exstruct/cli/edit.py`.

### Verification plan

- `tests/cli/test_cli_lazy_imports.py`
  - `import exstruct.cli.edit` does not eagerly load `pydantic`
  - `main(["--help"])` does not import `exstruct.cli.edit`
  - `typing.get_type_hints(exstruct.extract)` resolves `WorkbookData` successfully
- `tests/cli/test_edit_cli.py`
  - existing edit CLI behavior still passes with the new explicit loaders
- `uv run pytest tests/cli/test_cli_lazy_imports.py tests/cli/test_edit_cli.py tests/cli/test_cli.py -q`
- `uv run task precommit-run`

### ADR verdict

- `not-needed`
- rationale: this is a follow-up implementation hardening and static-analysis cleanup under the existing issue `#108` design, not a new policy decision.

## 2026-03-21 issue #108 review follow-up: validate runtime error scope

### Goal

- Restore the original `validate` subcommand exception boundary after the lazy-loader refactor in `src/exstruct/cli/edit.py`.
- Keep the patch/make commands catching `RuntimeError` while ensuring `validate` does not silently absorb it.

### Public contract

- `patch` and `make` continue to convert backend/runtime failures in `(OSError, RuntimeError, ValidationError, ValueError)` into `Error: ...` stderr output with exit code `1`.
- `validate` keeps its narrower historical contract and only converts `(OSError, ValidationError, ValueError)` into CLI error output.
- If `validate_input(...)` raises `RuntimeError`, the exception must still propagate rather than being turned into a handled CLI error.

### Constraints

- Do not broaden this follow-up into another startup optimization pass.
- Keep the current lazy import boundary for `pydantic` and validation helpers intact.
- Do not change the behavior of `patch` and `make` while narrowing `validate`.

### Verification plan

- `tests/cli/test_edit_cli.py`
  - `validate` still returns handled CLI errors for `OSError`
  - `validate` propagates `RuntimeError`
- `uv run pytest tests/cli/test_edit_cli.py -q`
- `uv run task precommit-run`

### ADR verdict

- `not-needed`
- rationale: this is a narrow behavior-restoration follow-up inside the existing edit CLI contract, not a new design decision.
