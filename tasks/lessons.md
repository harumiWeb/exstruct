## 2026-02-27 Review Fix Lessons

- When introducing structured error wrappers (`PatchOpError`), re-check outer fallback branches (`backend=auto`) so resilience paths are not accidentally bypassed.
- For `failed_field` inference from message text, avoid single hard-coded mapping for shared phrases like `sheet not found`; infer from contextual tokens (`category`) when available.
