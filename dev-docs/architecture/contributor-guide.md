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
