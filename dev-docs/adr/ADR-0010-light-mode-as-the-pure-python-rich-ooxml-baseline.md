# ADR-0010: Light Mode as the Pure-Python Rich OOXML Baseline

## Status

`proposed`

## Background

ExStruct currently treats `light` as the minimal extraction mode: cells, table
candidate detection, and print areas without any shapes or charts. That
boundary is documented in public docs, internal specs, engine docstrings, and
tests.

That contract no longer matches the intended product direction for non-COM
environments. The next target is not to redefine the existing LibreOffice path
as "no longer LibreOffice-dependent"; it is to make environments that only have
the current `light`-level capability return useful diagram and chart data
through pure Python alone.

The repository already contains part of that capability:

- `src/exstruct/core/ooxml_drawing.py` parses worksheet OOXML drawing parts into
  shapes, connectors, and chart metadata.
- `src/exstruct/core/backends/libreoffice_backend.py` already contains an
  OOXML-only shape/connector builder and already builds chart metadata from
  OOXML before applying LibreOffice geometry refinement.

What is missing is the product decision for mode boundaries and public contract
surface:

- whether `light` itself should start returning rich OOXML artifacts
- what remains unique to `libreoffice`
- how backend metadata should identify pure-Python rich artifacts when
  `include_backend_metadata=True`

Non-goals for this decision:

- matching COM fidelity for grouped shapes, SmartArt, or layout semantics
- adding pure-Python rich extraction for `.xls`
- changing PDF/PNG rendering or auto page-break policy

## Decision

- For `.xlsx` and `.xlsm`, `light` becomes the baseline pure-Python rich
  extraction mode.
- `light` remains free of Excel COM and LibreOffice runtime dependencies.
- `light` continues to include the current cell-first artifacts:
  - rows
  - table candidates
  - print areas
- In addition, `light` may emit best-effort OOXML-derived rich artifacts when
  present:
  - shapes
  - connectors
  - charts
- `light` rich artifacts are explicitly best-effort rather than COM-equivalent:
  - shape/chart metadata may be partial
  - artifacts unsupported by the current OOXML parser may be omitted
  - grouped-shape expansion and SmartArt reconstruction are not promoted to
    required `light` guarantees by this ADR
- `.xls` workbooks do not gain the new rich behavior; `light` remains minimal
  there because the OOXML drawing path is not available.
- `libreoffice` remains an opt-in non-COM enrichment mode for `.xlsx/.xlsm`.
  Its role shifts from "the only non-COM rich mode" to "an optional
  higher-fidelity refinement layer" that may improve geometry, emitted order, or
  connector matching beyond the `light` OOXML baseline.
- When `mode="libreoffice"` is requested for an OOXML workbook and the
  LibreOffice runtime is unavailable, ExStruct should preserve the `light`
  pure-Python rich baseline instead of dropping all the way to cells + table
  candidates only.
- `standard` and `verbose` remain COM-oriented modes with higher-fidelity native
  Excel extraction and are unchanged by this ADR.
- When backend metadata is requested, pure-Python rich artifacts use
  `provenance="python_ooxml"`. Existing `excel_com` and `libreoffice_uno`
  provenance values remain unchanged for their respective backends.

## Consequences

- Non-COM and non-LibreOffice hosts gain practical access to shapes,
  connectors, and charts for OOXML workbooks without requiring a secondary
  runtime.
- `light` becomes materially more useful in production, but it is no longer the
  smallest-possible extraction contract; payload size and token cost can
  increase on workbooks that contain many drawing objects.
- Existing users or tests that rely on `light` having empty `shapes` and
  `charts` must be updated explicitly. This is a public contract change, not an
  internal optimization.
- The semantic gap between `light` and `libreoffice` narrows. Future docs and
  guidance must explain `libreoffice` as an enrichment path rather than the only
  non-COM route to rich artifacts.
- The new `python_ooxml` provenance literal expands the public backend-metadata
  schema and requires synchronized updates across models, schemas, docs, and
  tests.
- Rich extraction fidelity in `light` still remains below COM and may remain
  below LibreOffice enrichment for geometry/order-sensitive workbooks. That
  trade-off is accepted in exchange for broader runtime availability.

## Rationale

- Tests:
  - `tests/core/test_mode_output.py`
  - `tests/integration/test_end_to_end_light.py`
  - `tests/core/test_libreoffice_backend.py`
  - `tests/core/test_libreoffice_smoke.py`
  - Additional `light`-mode regression tests for rich OOXML output do not exist
    yet and are required before this ADR can move from `proposed` to
    `accepted`.
- Code:
  - `src/exstruct/core/ooxml_drawing.py`
  - `src/exstruct/core/backends/libreoffice_backend.py`
  - `src/exstruct/core/pipeline.py`
  - `src/exstruct/models/__init__.py`
  - `src/exstruct/__init__.py`
  - `src/exstruct/engine.py`
- Related specs:
  - `dev-docs/specs/excel-extraction.md`
  - `dev-docs/specs/data-model.md`
  - `docs/cli.md`
  - `docs/api.md`
  - `docs/mcp.md`

## Supersedes

- `ADR-0001`

## Superseded by

- None
