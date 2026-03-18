## 2026-02-27 Review Fix Lessons

- When introducing structured error wrappers (`PatchOpError`), re-check outer fallback branches (`backend=auto`) so resilience paths are not accidentally bypassed.
- For `failed_field` inference from message text, avoid single hard-coded mapping for shared phrases like `sheet not found`; infer from contextual tokens (`category`) when available.

## 2026-02-27 apply_table_style COM compatibility lessons

- If helper functions are introduced to normalize API variants (property/callable), remove or relax pre-checks at call sites that can short-circuit before the helper runs.
- For COM compatibility fixes, add at least one integration-adjacent unit test at the higher-level caller path, not only helper-level tests.

## 2026-02-28 spec/implementation alignment lessons

- When a feature spec changes policy (for example, mixed-op allow/deny), re-run a direct implementation-vs-spec check before reporting completion.
- For policy flips, add one positive test and one negative boundary test in the same change so behavior drift is detected immediately.

## 2026-02-28 capture_sheet_images timeout hardening lessons

- For long-running COM/render paths exposed via MCP, always add an explicit tool-level timeout so client-side disconnects do not look like random transport failures.
- In multiprocessing paths, do not use unbounded `join()`; enforce `join(timeout)` and terminate/kill fallback to avoid hung workers blocking COM cleanup.

## 2026-02-28 non-finite timeout validation lessons

- For any env-based timeout parsed via `float()`, always reject non-finite values (`NaN`, `inf`, `-inf`) with `math.isfinite(...)` before range checks.
- When adding timeout hardening, include explicit regression tests for `NaN/inf/-inf`; testing only invalid strings and `<= 0` is insufficient.

## 2026-03-03 subprocess wait-order regression lessons

- In multi-stage timeout flows, define one primary end-to-end budget explicitly (here: join timeout) and ensure secondary timeouts are only local grace windows.
- When changing wait order, add regression tests that exceed the secondary timeout while staying inside the primary timeout to prevent accidental global timeout shrinkage.

## 2026-03-03 capture evaluation modal-dialog lessons

- For unattended Excel render evaluations, do not use fixed `A1:A1` as the minimal-range case; select a known non-empty single cell per workbook.
- Add a run-validity rule for Excel modal dialogs (invalid run + rerun), otherwise stability metrics can be overstated.
- In render paths that open Excel for export, explicitly set `app.display_alerts = False` even if other paths already do so.

## 2026-03-06 libreoffice/ooxml review lessons

- When mapping OOXML connector semantics into internal arrow fields, verify `head`/`tail` against the source spec instead of inferring from names alone; add separate start/end regression tests.
- If `__enter__` allocates temp resources before a subprocess probe, clean them up in the exception path as well; `__exit__` is not guaranteed to run on enter failure.

## 2026-03-06 libreoffice validator contract lessons

- When composing higher-level validators from lower-level ones, keep each validator sound on its own contract; do not suppress a lower-level check just to improve a combined error path unless the caller fully re-implements that check.
- If a validator has branching for combined invalid options, add a direct unit test for the single-option branch and the combined branch so downstream callers do not mask a contract hole.

## 2026-03-06 docs parity lessons

- When changing a public README example or CLI/API option in `README.md`, update `README.ja.md` in the same change before reporting completion.
- For token/serialization policy changes, check both English and Japanese quick-start sections for parity on defaults and opt-in flags.


## 2026-03-10 libreoffice smoke gate retry lessons

- For Windows cold-start runtime checks, avoid single-shot `soffice --version` gating with a short timeout; add an explicit longer retry before declaring runtime unavailable.
- If a fallback probe is expensive (full session startup), place a cheaper retry tier ahead of it to reduce false negatives under CI install jitter.

## 2026-03-13 ADR governance contract alignment lessons

- When a shared policy document defines a required output artifact (here: the `specs`/`src`/`tests` evidence triad), mirror that requirement in every dependent skill contract; do not assume downstream docs will fill the gap.
- In decision workflows, collect verification evidence before any terminal verdict, including negative outcomes like `not-needed`; otherwise the process silently permits ungrounded dismissals.

## 2026-03-13 ADR index contract lessons

- When a spec makes structured audit fields mandatory (for example `scope` or finding `type`), copy those exact fields into the producing skill contract; partial paraphrases in workflow docs are not enough.
- If a human-facing artifact needs one canonical label while machine-readable metadata supports multiple labels, encode the canonical label explicitly (for example `primary_domain`) instead of inferring it from array order or merged headings.

## 2026-03-13 ADR reviewer scope and gating lessons

- When a review skill is responsible for compatibility or public break judgment, make the relevant public `docs/` pages part of its required scope; internal specs alone are not enough evidence.
- When lint and design review are split into separate skills, encode a clean linter result as an explicit precondition in the skill, spec, and workflow docs so `ready` cannot bypass mandatory structural checks.

## 2026-03-13 AGENTS retention policy lessons

- When AGENTS explains how to preserve or migrate durable documentation, explicitly direct agents to the relevant repository skills; otherwise the ADR workflow is easy to bypass with ad hoc manual judgment.

## 2026-03-13 ADR review follow-up lessons

- When a template defines a required section such as `状態`, mirror that exact requirement in the producing or linting skill checklist; validating only the value is not enough if the section itself can be omitted.
- When recording validation commands in tracked docs, avoid machine-specific absolute paths; use a placeholder or portable form so the evidence remains reproducible across contributors.

## 2026-03-16 pytest collection naming lessons

- When adding new pytest files under different directories, keep the basename unique across the repository unless the directories are explicit Python packages; duplicate `test_*.py` basenames can trigger `import file mismatch` during collection.
- For new test modules, run a targeted `--collect-only` check against any similarly named legacy test files before reporting completion.

## 2026-03-16 compatibility shim monkeypatch lessons

- When preserving legacy monkeypatch surfaces, do not forward compatibility wrappers through copied function aliases; use live module lookup so monkeypatches on legacy modules remain observable at call time.
- If one compatibility entrypoint re-synchronizes another layer before execution, add a regression test for override precedence at the highest public entrypoint; function identity checks alone are insufficient.

## 2026-03-18 docs positioning lessons

- When ExStruct exposes a Python wrapper over behavior that existing ecosystem libraries already handle well, do not automatically promote that wrapper as the default Python recommendation; confirm the intended positioning first.
- For workbook editing docs, bias the primary recommendation toward the editing CLI for ExStruct-specific workflows and keep `exstruct.edit` described as an advanced/shared-contract surface unless the user explicitly wants stronger promotion.

## 2026-03-18 docs review follow-up lessons

- When documenting a `dry_run -> apply` edit workflow, do not imply the same engine will run both phases under `backend="auto"`; call out the openpyxl/COM split and tell users to pin `openpyxl` when same-engine comparison matters.
- When documenting CLI failure behavior, distinguish serialized `PatchResult.error` failures from pre-execution stderr failures such as JSON parse, validation, or local runtime errors.
