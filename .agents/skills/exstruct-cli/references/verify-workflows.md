# Verify Workflows

Use this file after `patch` or `make`, or after a `--dry-run` when the result
must be inspected before apply.

## Always check

- CLI exit code
- `PatchResult.error`
- `PatchResult.warnings`
- `PatchResult.engine`
- `PatchResult.patch_diff`
- `PatchResult.out_path`

## After `--dry-run`

- Confirm that `patch_diff` matches the intended change.
- Confirm that warnings and formula issues are acceptable.
- If the user needs same-engine confidence, ensure the command pinned
  `--backend openpyxl`.

## After real apply

- Re-run `exstruct validate` when the workbook was unfamiliar or the change was
  operationally sensitive.
- Re-run extraction or another downstream read step when the workbook feeds a
  larger pipeline and the user needs output-level confirmation.
- Prefer lightweight verification for small, low-risk edits and deeper
  verification for wide-range, formula, chart, or backend-sensitive changes.

## When lightweight verification is enough

- The change is narrow and the `patch_diff` is easy to inspect.
- No backend-sensitive features were used.
- The user did not request stronger confirmation or downstream artifacts.
