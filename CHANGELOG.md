# Changelog

All notable changes to this project are documented in this file. This changelog follows the [Keep a Changelog](https://keepachangelog.com/) format and covers changes from v0.2.70 onward.

## [Unreleased]

## [0.8.0] - 2026-04-22

### Added

- Added typed LibreOffice workbook handles and session-scoped workbook lifecycle tracking so rich extraction can reuse cached bridge payloads safely while rejecting foreign or closed workbook handles.
- Added a pure-Python OOXML rich backend for `.xlsx` / `.xlsm`, allowing `mode="light"` to emit best-effort shapes, connectors, and charts without Excel COM or LibreOffice.
- Added regression coverage for light-mode OOXML rich extraction, print-area side outputs, per-sheet OOXML drawing failures, and LibreOffice baseline/enrichment fallback behavior.

### Changed

- Changed the `light` mode contract so public API, engine, and CLI paths keep print areas by default and expose the OOXML-rich baseline consistently across `extract`, `process_excel`, and CLI output paths.
- Changed `libreoffice` mode to seed the same OOXML baseline before optional UNO enrichment so non-COM fallback can preserve already recovered rich artifacts.
- Updated ADR/spec/docs/schema artifacts to describe `light` as the pure-Python OOXML-rich baseline and to expose `python_ooxml` backend metadata in serialized models.

### Fixed

- Fixed LibreOffice rich backend workbook lifecycle integration so custom `session_factory` implementations that only support legacy path-based `extract_chart_geometries()` and `extract_draw_page_shapes()` continue to work without `load_workbook()` and `close_workbook()` hooks.
- Fixed OOXML drawing resilience so malformed or corrupt worksheet drawing parts only skip the affected sheet instead of clearing healthy workbook siblings.
- Fixed `process_excel()` and engine filter alignment so `FilterOptions.include_print_areas=None` once again means automatic inclusion instead of an implicit hard-coded override.
- Fixed light/libreoffice review follow-up edge cases by hardening OOXML baseline seeding, streaming worksheet metrics reads, caching cumulative row/column offsets, and correcting stale README / architecture wording.

## [0.7.1] - 2026-03-21

### Added

- Added regression coverage for extraction CLI runtime validation and lightweight import boundaries across `exstruct`, `exstruct.engine`, `exstruct.cli.main`, and `exstruct.cli.edit`.

### Changed

- Changed the extraction CLI so `--auto-page-breaks-dir` is always listed in help output and validated only when the flag is requested at runtime.
- Changed CLI and package import behavior so `exstruct --help`, `exstruct ops list`, `import exstruct`, and `import exstruct.engine` defer heavy extraction, edit, and rendering imports until needed.

### Fixed

- Fixed parser and help startup side effects by removing COM availability probing during extraction CLI parser construction.
- Fixed lazy-export follow-ups so public runtime type hints resolve correctly while keeping exported symbol names stable.
- Fixed edit CLI routing so non-edit argv and lightweight edit paths avoid unnecessary imports such as `exstruct.cli.edit` and `pydantic`.
- Fixed the `validate` subcommand error boundary so `RuntimeError` is no longer converted into handled CLI stderr output.

## [0.7.0] - 2026-03-19

### Added

- Added a first-class public workbook editing API under `exstruct.edit`, including public patch/make entrypoints, shared patch-op schema helpers, and edit-owned request/result models.
- Added public editing CLI commands under the existing `exstruct` console script: `patch`, `make`, `ops`, and `validate`.
- Added maintainer-facing editing documentation coverage, including architecture/spec updates, ADR alignment, and agent workflow guidance that closes out issue `#99`.

### Changed

- Changed workbook editing layering so `exstruct.edit` is the canonical editing core while MCP remains a host-managed integration and compatibility layer.
- Updated README and docs positioning to clarify canonical usage across Python, CLI, and MCP workflows, including dry-run guidance for editing operations.

### Fixed

- Fixed top-level `sheet` fallback handling for workbook editing requests while preserving `op.sheet` precedence.
- Fixed legacy monkeypatch compatibility across `exstruct.mcp.patch_runner` and related compatibility shims by restoring live override visibility and entrypoint precedence coverage.
- Fixed rename-reservation cleanup on openpyxl failure paths so placeholder output files are removed when apply fails.
- Fixed dry-run, backend-selection, and CLI failure wording drift in the docs so it matches current runtime behavior.

## [0.6.1] - 2026-03-12

### Added

- Added a dedicated GitHub Actions Windows LibreOffice smoke job on `windows-2025` that installs `libreoffice-fresh`, discovers runtime paths, and runs `tests/core/test_libreoffice_smoke.py` with `RUN_LIBREOFFICE_SMOKE=1`.
- Added Windows-focused regression coverage for LibreOffice runtime normalization, bundled Python discovery, bridge subprocess environment setup, and smoke-gate timeout fallback behavior.

### Changed

- Updated README, README.ja, and test requirements to document LibreOffice smoke coverage on both Linux and Windows CI.
- Changed LibreOffice bridge subprocess execution on Windows so probe, handshake, and extraction runs use the runtime directory as `cwd` and prepend runtime paths to `PATH`.

### Fixed

- Fixed Windows LibreOffice runtime discovery to prefer `soffice.com` when it is available and to detect bundled LibreOffice Python under `python-core-*` layouts.
- Fixed false-negative Windows LibreOffice smoke gating by retrying slow `soffice --version` probes and falling back to a short-lived session probe before treating the runtime as unavailable.

## [0.6.0] - 2026-03-06

### Added

- Added a new `libreoffice` extraction mode across the Python API, CLI, and MCP. This mode provides best-effort rich extraction for `.xlsx`/`.xlsm` without Excel COM and can add merged cells, shapes, connectors, and charts when the LibreOffice runtime is available.
- Added a LibreOffice-backed rich extraction pipeline, including headless session management, timeout/profile cleanup handling, explicit fallback reasons, and non-COM fallback workbook generation when the runtime is unavailable.
- Added best-effort shape, connector, and chart reconstruction for `libreoffice` mode by combining LibreOffice UNO geometry with OOXML metadata.
- Added provenance/fidelity metadata for rich objects: shapes and charts now report `provenance`, `approximation_level`, and `confidence`.
- Added LibreOffice-focused regression coverage, including mode validation, `.xls` rejection, connector/chart extraction, unavailable-runtime fallback, and optional smoke tests.

### Changed

- Updated docs across README, CLI, API, architecture, and release notes to describe `libreoffice` as a best-effort rich mode rather than a strict subset of COM output.
- Updated pipeline/backend reporting so `light`, `libreoffice`, and COM-backed rich extraction paths are distinguished more clearly.
- Clarified public contracts and help text for mode support, fallback behavior, and LibreOffice limitations in v1.

### Fixed

- Fixed early validation for `mode="libreoffice"` so unsupported combinations with PDF/PNG rendering and auto page-break export now fail consistently in CLI and API before processing starts.
- Fixed unsupported `.xls` handling in `libreoffice` mode by returning a clear early error instead of attempting runtime processing.

## [0.5.3] - 2026-03-03

### Added

- Added a dedicated render worker entrypoint (`python -m exstruct.render.subprocess_worker`) for `capture_sheet_images` subprocess mode, decoupled from parent `__main__` restoration.

### Changed

- MCP runtime now defaults `EXSTRUCT_RENDER_SUBPROCESS=1` after profile comparison runs showed stable behavior in both modes (`63/63` success for `0` and `1` under MCP-equivalent timeout handling); set `EXSTRUCT_RENDER_SUBPROCESS=0` to force in-process rendering.
- Marked MCP `exstruct_capture_sheet_images` as Experimental in docs, including recommended timeout/runtime settings.
- Updated MCP/README docs with subprocess timeout tuning and stage-aware error guidance (`startup`/`join`/`result`/`worker`), including `EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC`.

### Fixed

- Fixed subprocess render wait ordering to prioritize result receipt before join wait, preventing false timeout failures after successful worker output.
- Fixed opaque subprocess failures by returning actionable stage-aware render errors with stderr snippets where available.

## [0.5.2] - 2026-02-28

### Fixed

- Restored support for mixed `create_chart` + `apply_table_style` requests in one run when backend resolves to COM (`backend="com"` or `backend="auto"` with COM available).
- Improved mixed-op error behavior when COM is unavailable by returning a clear COM-required message for `create_chart` + `apply_table_style` requests.

### Changed

- Updated MCP/README docs to reflect mixed chart+table request support and backend requirements.

## [0.5.1] - 2026-02-26

### Added

- Added explicit service-level guard for mixed backend-only patch ops:
  `create_chart` and `apply_table_style` can no longer be combined in one request.

### Changed

- Updated MCP docs and README pages to document `create_chart` backend constraints
  (COM-only, flag limitations, and incompatibility with `apply_table_style` in one request).

## [0.5.0] - 2026-02-24

### Added

- Added MCP `exstruct_make` for one-call workbook creation plus `ops` apply (`out_path` required, `ops` optional), including `.xlsx`/`.xlsm`/`.xls` support and `.xls` COM constraints.
- Expanded MCP `exstruct_patch` with design editing operations: `draw_grid_border`, `set_bold`, `set_font_size`, `set_font_color`, `set_fill_color`, `set_dimensions`, `auto_fit_columns`, `merge_cells`, `unmerge_cells`, `set_alignment`, `set_style`, `apply_table_style`, and inverse restore op `restore_design_snapshot`.
- Added MCP operation schema discovery tools: `exstruct_list_ops` and `exstruct_describe_op`.
- Added MCP runtime diagnostics tool: `exstruct_get_runtime_info`.
- Added top-level `sheet` fallback for `exstruct_patch`/`exstruct_make` (non-`add_sheet` ops), with `op.sheet` precedence when both are provided.
- Added artifact mirroring support via `mirror_artifact` and server `--artifact-bridge-dir`.

### Changed

- Updated patch backend controls for MCP `exstruct_patch`/`exstruct_make`: `backend` input (`auto`/`com`/`openpyxl`) and `engine` output (`com`/`openpyxl`).
- Updated patch backend policy: `auto` now prefers COM when available, with controlled fallback to openpyxl for `.xlsx`/`.xlsm` when COM execution fails.
- Updated `apply_table_style` behavior: when `backend="com"` is requested, execution falls back to openpyxl with a warning.
- Refactored MCP patch internals into layered modules (`patch.service` / `patch.engine.*` / `patch.ops.*` / `patch.runtime`) while keeping tool interfaces stable.
- Updated MCP docs/README pages to include `exstruct_make` behavior and constraints.

## [0.4.4] - 2026-02-16

### Added

- Added an MVP of Excel editing for MCP via `exstruct_patch`, including atomic apply semantics and expanded operations: `set_range_values`, `fill_formula`, `set_value_if`, and `set_formula_if`.
- Added direct A1-oriented MCP read tools for extracted JSON: `exstruct_read_range`, `exstruct_read_cells`, and `exstruct_read_formulas`.
- Added patch safety/review options: `dry_run`, `return_inverse_ops`, `preflight_formula_check`, and `auto_formula`.

### Changed

- Improved `exstruct_patch` input compatibility: `ops` now accepts both object lists (recommended) and JSON object strings.
- Enabled `alpha_col` support more broadly across extraction/read flows, and added `merged_ranges` output support for alpha-column mode.
- Updated MCP documentation and chunking guidance, including clearer error messages and mode guidance.
- Changed MCP default conflict policy to `overwrite` for output handling.

## [0.4.2] - 2026-01-23

### Changed

- Renamed MCP tool names to remove dots for compatibility with strict client validators (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).

## [0.4.1] - 2026-01-23

### Fixed

- Pinned `httpx<1.0` for MCP extras to prevent runtime failures with pre-release `httpx` builds (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).

## [0.4.0] - 2026-01-23

### Added

- Added a stdio MCP server (`exstruct-mcp`) with tool discovery and invocation (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).
- Added MCP tools: `exstruct_extract`, `exstruct_read_json_chunk`, and `exstruct_validate_input` (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).
- Added MCP `exstruct[mcp]` extras with required dependencies, plus documentation and examples for agent configuration (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).
- Added MCP safety controls: root allowlist enforcement, deny-glob support, and conflict handling (`--on-conflict`) (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).

### Fixed

- Pinned MCP HTTP client dependency to stable `httpx<1.0` to avoid runtime errors in MCP initialization (PR [#47](https://github.com/harumiWeb/exstruct/pull/47)).

## [0.3.7] - 2026-01-23

### Added

- Added formula extraction via a new `formulas_map` output field (maps formula strings to cell coordinates), enabled by default in **verbose** mode (PR [#44](https://github.com/harumiWeb/exstruct/pull/44)).

### Fixed

- Improved print-area exports to be more robust: all print areas are now numbered safely and errors during print area restoration are handled gracefully, ensuring no missing pages or crashes.

## [0.3.6] - 2026-01-12

### Added

- Added an option to run Excel rendering in a separate subprocess (enabled by default) to improve stability on large workbooks. This isolates memory usage during PDF/PNG generation. Set `EXSTRUCT_RENDER_SUBPROCESS=0` to disable this behavior if needed (PR [#41](https://github.com/harumiWeb/exstruct/pull/41)).

### Fixed

- Fixed sheet image exports for multi-page print ranges: previously only the first page image was output; now all pages are exported with suffixes `_pNN` for page 2 and beyond (PR [#41](https://github.com/harumiWeb/exstruct/pull/41)).
- Fixed image exports for legacy `.xls` files by automatically converting them to `.xlsx` via Excel before rendering. This prevents failures when exporting images from older Excel formats (PR [#41](https://github.com/harumiWeb/exstruct/pull/41)).

## [0.3.5] - 2026-01-06

### Breaking Changes

- The JSON structure for `merged_cells` in outputs has changed (PR [#40](https://github.com/harumiWeb/exstruct/pull/40)). In versions <= 0.3.2, `merged_cells` was an array of objects; in v0.3.5 it is now an object with a `schema` definition and `items` list of merged cell ranges.

### Migration Guide

- If upgrading from an older version, update any code that parses `merged_cells`. Expect an object with `schema` and `items` instead of a simple list. Refer to the updated README for detailed transition guidance on the new format.

### Added

- Added a configuration flag `include_merged_values_in_rows` in `StructOptions` to control whether values from merged cells are duplicated in the main `rows` output. This flag defaults to **True** for backward compatibility (PR [#40](https://github.com/harumiWeb/exstruct/pull/40)).

### Changed

- `merged_cells` output format now uses a compact schema-based structure (see Breaking Changes above).
- Empty merged cells (merged ranges with no content) are now represented as a single space `" "` in the output, to clearly denote an intentional blank (PR [#40](https://github.com/harumiWeb/exstruct/pull/40)).

## [0.3.2] - 2026-01-05

### Added

- Added extraction of merged cell ranges. Each sheet's output now includes a `merged_cells` field listing all merged cell ranges with their coordinates (PR [#35](https://github.com/harumiWeb/exstruct/pull/35)).
- Added options to control merged cell output: you can disable including merged cells via `StructOptions.include_merged_cells` or `OutputOptions.filters.include_merged_cells` if you do not want this data in the output (PR [#35](https://github.com/harumiWeb/exstruct/pull/35)).

### Changed

- Standard and verbose mode outputs now include `merged_cells` by default (PR [#35](https://github.com/harumiWeb/exstruct/pull/35)). If your workflow does not need merged cell information, use the provided options to omit it.

## [0.3.1] - 2025-12-28

### Breaking Changes

- The shape output format has changed to accommodate SmartArt extraction. SmartArt shapes now use a new nested node structure and some previously existing fields have been removed or renamed:
  - Removed output fields `layout_name`, `roots`, and `children` for SmartArt. These are replaced by a new `layout` field and a nested `nodes` list (with child nodes under `kids`).
  - The `type` field is no longer present on Arrow (connector) and SmartArt shape outputs (it remains only for regular shape types).

### Migration Guide

- Update any code that parses shape outputs, especially for SmartArt diagrams. Instead of `layout_name` and nested `children`, use the new `layout` and `nodes` (with `kids`) format for SmartArt. Arrow and SmartArt objects will not include a `type` field anymore, so ensure your code doesn’t assume its presence.

### Added

- Added **SmartArt extraction** support (Excel COM required). SmartArt diagrams in Excel are now parsed and included in the output, with each SmartArt represented by a `kind: "smartart"` shape containing a `layout` name and a hierarchical `nodes` structure of text entries.
- The shape model now differentiates between regular shapes, connectors (arrows), and SmartArt, providing clearer semantics in the output JSON.

### Changed

- Internal shape handling has been refactored to support SmartArt: shapes of `kind: "arrow"` (connectors) and `kind: "smartart"` are now separate from standard shapes, each with their appropriate fields. This improves clarity but may require the adjustments noted in the Migration Guide.

## [0.3.0] - 2025-12-27

### Changed

- Major **internal refactor** of the processing pipeline and code structure to improve maintainability and enable future features (PR [#23](https://github.com/harumiWeb/exstruct/pull/23)). There are **no user-facing API changes** or behavior changes in this release.

## [0.2.90] - 2025-12-24

### Added

- Added extraction of cell background colors via a new `colors_map` field in each sheet’s output. The `colors_map` maps color hex codes to lists of cell coordinates that have that background color. In Excel COM environments, this includes evaluation of conditional formatting colors (PR [#21](https://github.com/harumiWeb/exstruct/pull/21)).
- Added `ColorsOptions` (e.g., `include_default_background` and `ignore_colors`) to allow configuration of color extraction. You can exclude default fill colors or ignore specific colors to reduce output size.

### Changed

- **Verbose** mode now enables `colors_map` by default, so detailed color information will be included unless explicitly disabled. Non-COM environments still extract static fill colors via openpyxl, but cannot detect conditional formats.

## [0.2.80] - 2025-12-21

### Added

- Added unique shape IDs for more robust flowchart tracing: each non-connector shape now receives a sequential `id` per sheet for stable reference in connectors.
- Connector (arrow) shapes now include references to their connected shapes: each connector output has `begin_id` and `end_id` fields pointing to the IDs of the shapes it connects (via Excel COM’s ConnectorFormat) (PR [#15](https://github.com/harumiWeb/exstruct/pull/15)).
- Added extra metadata for connectors such as arrow style, direction, and rotation in the output JSON, to enrich flowchart and diagram analysis.

## [0.2.71] - 2025-12-17

### Added

- Added CLI support for exporting **auto page-break** views. A new option `--auto-page-breaks-dir` allows saving each worksheet’s automatic page-break layout to separate files (when running on a system with Excel COM available).
- Documentation and help text have been updated to describe the new option, and tests were added to ensure it only appears when supported.

### Changed

- The CLI now dynamically detects Excel/COM availability and will only register COM-specific flags (such as `--auto-page-breaks-dir`) when Excel is usable. This prevents showing or using unsupported options on environments where Excel is not available.

## [0.2.70] - 2025-12-15

### Added

- Added more flexible file path handling: you can now pass file paths as simple `str` strings in addition to `pathlib.Path` objects for all engine inputs and outputs. All paths (including those for PDF/PNG rendering) are internally normalized to `Path` for consistent behavior.

### Changed

- Changed export behavior when only "secondary" outputs are requested. If you call the export function with `output_path=None` and specify only auxiliary directories (such as `sheets_dir`, `print_areas_dir`, or `auto_page_breaks_dir`), the tool will **no longer write to standard output** by default. It will only produce the specified secondary output files.

### Migration Guide

- If you need the combined output on stdout (as previous versions would do by default), make sure to provide an explicit `output_path` or use a `stream` in the export options. This will ensure that the main output is still sent to standard output when using secondary output directories.
