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
