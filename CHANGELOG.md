# Changelog

All notable changes to this project are documented in this file. This changelog follows the [Keep a Changelog](https://keepachangelog.com/) format and covers changes from v0.2.70 onward.

## [Unreleased]

### Added

- _No unreleased changes yet._

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
