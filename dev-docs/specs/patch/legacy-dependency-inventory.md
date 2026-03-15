# Legacy Dependency Inventory (Phase 2)

Updated: 2026-02-24

`src/exstruct/mcp/patch/legacy_runner.py` was deleted upon Phase 2 completion.
This document records the inventory results for the old dependencies and their current replacements.

## Replacement Targets for Old Dependencies

- Old target: `src/exstruct/mcp/patch/legacy_runner.py` (deleted)
- Current responsibility split:
  - `src/exstruct/mcp/patch/service.py`: patch/make orchestration
  - `src/exstruct/mcp/patch/engine/openpyxl_engine.py`: openpyxl backend execution boundary
  - `src/exstruct/mcp/patch/engine/xlwings_engine.py`: COM (xlwings) backend execution boundary
  - `src/exstruct/mcp/patch/runtime.py`: runtime utilities (engine selection, path, policy)
  - `src/exstruct/mcp/patch/ops/*`: backend-specific op application logic

## Compatibility Layer

- `src/exstruct/mcp/patch_runner.py`
  - Thin facade maintaining public import compatibility
  - Delegates actual implementation to `patch/service.py`

## Test Scope

- `tests/mcp/test_patch_runner.py`
  - Verifies the compatible entry point of `patch_runner` (delegation behavior)
- `tests/mcp/patch/test_service.py`
  - Verifies backend selection, fallback, and warning messages
