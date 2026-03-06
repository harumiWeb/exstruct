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
