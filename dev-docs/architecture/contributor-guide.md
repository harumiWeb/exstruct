# Contributor Guide — Internal Architecture

## Target Audience

This page is for people who:

- Want to extend ExStruct's internal implementation
- Want to add new extraction targets (shapes, SmartArt, comments, etc.)
- Want to extend a backend (Openpyxl / COM / LibreOffice / future XML)
- Are trying to submit a PR but are unsure which files to touch

---

## Directory Structure (core)

```text
src/exstruct/core/
├── pipeline.py        # Orchestrates the overall flow
├── backends/          # Backend abstractions and runtime-specific adapters
│   ├── openpyxl_backend.py
│   ├── com_backend.py
│   └── libreoffice_backend.py
├── libreoffice.py     # LibreOffice runtime/session helper
├── ooxml_drawing.py   # OOXML drawing/chart parser for best-effort rich extraction
├── modeling.py        # Final data integration
├── workbook.py        # Workbook lifecycle management
├── cells.py           # Cell/table analysis (mainly openpyxl)
└── utils.py           # Shared utilities
```

---

## Important Design Rules

### 1. Pipeline only knows the order

- Do not put Excel parsing logic in Pipeline
- Limit Pipeline's responsibilities to only the following:
  - Calling order of backends
  - Fallback decisions
  - Artifact management
  - Handoff to Modeling

**Decision criterion**

> Is this code directly reading Excel content?
> If so, it should not be in Pipeline.

---

### 2. Backend is for extraction only

Backend exists for **pure extraction**.

- Excel → raw data
- No interpretation
- No integration
- Avoid side effects as much as possible

#### What is allowed in Backend

- Reading cell values
- Reading shape positions
- Calling COM APIs
- Raising exceptions

#### What is not allowed in Backend

- Building WorkbookData / SheetData
- Bringing in concerns about the output format
- Fallback logging (this is Pipeline's responsibility)

---

### 3. Make Modeling the single integration point

Only Modeling should integrate results from multiple backends into a single **semantic structure**.

- Combine Openpyxl + COM / LibreOffice results
- Normalize coordinates, directions, and types
- Fill in missing data

> The only layer that may know the final JSON/YAML/TOON shape
> is **Modeling**.

---

## Common Extension Patterns

---

## Case 1: Adding a New Extraction Target (e.g., comments)

### Steps

1. **Add an extraction method to Backend**

   ```python
   class Backend(Protocol):
       def extract_comments(self, ...): ...
   ```

2. Implement in `OpenpyxlBackend` / `ComBackend`
   - One side is enough. Use `NotImplementedError` if not implemented.

3. Add the call to `pipeline.py`
   - Explicitly state whether to include it as a fallback target.

4. Integrate into WorkbookData in `modeling.py`

5. Add tests

---

## Case 2: Adding a New Backend (e.g., XML or LibreOffice backend)

### Steps

1. Implement `Backend` and/or `RichBackend` from `src/exstruct/core/backends/base.py` in a new backend module

   ```python
   class XmlBackend:
        def extract_cells(self, *, include_links: bool):
            ...

        def extract_shapes(self, *, mode: str):
            ...
   ```

2. Add backend selection to Pipeline
   - Minimize changes to existing backends.

3. Keep Modeling unchanged if possible

---

## Case 3: Changing the Output Structure

- **This is the most fragile type of change**

### Principles

- Limit changes to `modeling.py` and the Pydantic model
- Do not change the backend
- Do not change Pipeline

---

## LibreOffice Session Contract

`LibreOfficeRichBackend` accepts a `session_factory` that returns a
context-managed rich-extraction session (default: `LibreOfficeSession.from_env`).
Custom integrations may supply either of two structural session contracts, and
the backend must keep supporting both:

- **Legacy path-only sessions**: expose `extract_chart_geometries(file_path)`
  and `extract_draw_page_shapes(file_path)`, each taking a workbook path.
- **Lifecycle-aware sessions**: additionally expose `load_workbook(file_path)`
  and `close_workbook(workbook)`; their extraction methods accept a path or a
  typed `LibreOfficeWorkbookHandle`.

Lifecycle support is detected structurally via `load_workbook` /
`close_workbook`. The path-only path is a supported extension contract for
custom `session_factory` integrations and must not be removed as dead code.

For lifecycle-aware sessions, `LibreOfficeSession.load_workbook()` returns a
frozen typed handle bound to the resolved workbook path and its owning session,
and `close_workbook()` validates that the handle belongs to the session,
rejects rehydrated handles whose `file_path` no longer matches the registered
workbook id, stays idempotent across repeated close calls, and clears
session-local bridge cache entries for that workbook.

These session contracts do not change the public CLI, MCP, extraction-mode,
fallback, or serialization contracts.

---

## Fallback Rules

- COM or LibreOffice runtime being unavailable is **the normal case**
- Do not treat fallback as an exception
- Always provide a `FallbackReason`

```python
log_fallback(
    reason=FallbackReason.COM_UNAVAILABLE,
    message="COM backend not available"
)

log_fallback(
    reason=FallbackReason.LIBREOFFICE_UNAVAILABLE,
    message="LibreOffice backend not available"
)
```

---

## Testing Guidelines

### Expected test granularity

| Layer    | Test focus           |
| -------- | -------------------- |
| Backend  | extraction correctness |
| Pipeline | fallback / branching |
| Modeling | integration logic    |

### Anti-patterns

- Fragile tests that depend heavily on a real Excel instance
- Massive tests that couple Backend and Modeling all at once

---

## Pre-PR Checklist

- [ ] No Excel parsing logic in Pipeline
- [ ] No interpretation logic in Backend
- [ ] Modeling is the single source of truth for the final structure
- [ ] Fallback reason is explicit
- [ ] Tests have been added
- [ ] If the public API changed, docs have been updated

---

## Common Anti-patterns

- Building WorkbookData inside Backend
- Calling openpyxl / xlwings directly from Pipeline
- Ad-hoc logic that "just handles it here"
- Catch-all exceptions with no fallback reason

---

## Summary of Design Philosophy

- Excel is **fragile**
- COM is **powerful but unstable**
- LLM/RAG requires **stable structure first**

Therefore,

> Separate responsibilities and localize failure points.
