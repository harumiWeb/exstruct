# Patch Model Migration Notes (Phase 2)

Notes on dependencies when consolidating `patch/models.py` as the canonical model.

## Current Coupling Points

- `patch/models.py` has canonical definitions for `PatchOp` / `PatchRequest` / `PatchResult` and similar, while `internal.py` still has duplicate definitions
- `patch_runner.py` / `service.py` / `runtime.py` / `ops/*` depend on the private implementation in `internal.py`
- As a result, having both lineages of `BaseModel` coexist causes type mismatches in both mypy and runtime validation

## Recommended Incremental Migration Steps

1. Remove the duplicate model definitions in `internal.py` and replace them with imports from `patch/models.py`
2. Move the model validation helper functions in `internal.py` (related to `PatchOp`) to the `models.py` side
3. Unify the type annotations and return types in `runtime.py` / `ops/*` / `service.py` to use `patch.models`
4. Keep the compatibility tests in `tests/mcp/test_patch_runner.py` intact while migrating `internal.py`-dependent tests to `tests/mcp/patch/*`
5. Finally, wind down the compatibility dependency on `internal.py` and lock `patch_runner.py` as the thin public API entry point

## Cautions

- Switching call sites while keeping the duplicate definitions in `internal.py` makes validation fail easily because different class hierarchies mix when constructing `PatchResult`
- The safe approach is to **consolidate the definition source first** and then swap out the call sites
