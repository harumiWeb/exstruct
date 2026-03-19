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
