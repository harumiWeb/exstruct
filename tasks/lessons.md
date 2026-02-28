## 2026-02-27 Review Fix Lessons

- When introducing structured error wrappers (`PatchOpError`), re-check outer fallback branches (`backend=auto`) so resilience paths are not accidentally bypassed.
- For `failed_field` inference from message text, avoid single hard-coded mapping for shared phrases like `sheet not found`; infer from contextual tokens (`category`) when available.

## 2026-02-27 apply_table_style COM compatibility lessons

- If helper functions are introduced to normalize API variants (property/callable), remove or relax pre-checks at call sites that can short-circuit before the helper runs.
- For COM compatibility fixes, add at least one integration-adjacent unit test at the higher-level caller path, not only helper-level tests.

## 2026-02-28 spec/implementation alignment lessons

- When a feature spec changes policy (for example, mixed-op allow/deny), re-run a direct implementation-vs-spec check before reporting completion.
- For policy flips, add one positive test and one negative boundary test in the same change so behavior drift is detected immediately.
