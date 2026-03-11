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
