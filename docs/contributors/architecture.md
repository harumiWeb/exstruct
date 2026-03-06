# Contributor Guide - Internal Architecture

## Audience

This page is intended for contributors who:

- want to extend ExStruct internals
- want to add new extraction targets (shapes, SmartArt, comments, etc.)
- want to extend backends (Openpyxl / COM / LibreOffice / future XML)
- want to submit a PR but are unsure where to make changes

---

## Directory Layout (core)

```text
src/exstruct/core/
├── pipeline.py        # Orchestrates the overall flow
├── backends/          # Backend abstractions and runtime-specific adapters
├── libreoffice.py     # LibreOffice runtime/session helper
├── ooxml_drawing.py   # OOXML drawing/chart parser for best-effort rich extraction
├── openpyxl_backend.py
├── com_backend.py
├── modeling.py        # Final data integration
├── workbook.py        # Workbook lifecycle management
├── cells.py           # Cell/table analysis (mainly openpyxl)
└── utils.py           # Shared utilities
```

---

## Key Design Rules (Must Read)

### 1. The Pipeline only knows the order

- Do not place Excel parsing logic in Pipeline
- Pipeline is responsible for:
  - Backend call order
  - fallback decisions
  - artifact management
  - passing to Modeling

**Rule of thumb**

> "Does this code read Excel contents directly?"
> If yes, it does not belong in Pipeline.

---

### 2. Backend is extraction-only

Backends are for **pure extraction**:

- Excel -> raw data
- no interpretation, no integration
- avoid side effects whenever possible

#### Allowed in Backend

- reading cell values
- reading shape positions
- calling COM APIs
- raising exceptions

#### Not allowed in Backend

- building WorkbookData / SheetData
- output-format concerns
- fallback logging (Pipeline handles it)

---

### 3. Modeling is the single integration point

Only Modeling merges backend results into one **semantic structure**.

- combine Openpyxl + COM / LibreOffice results
- normalize coordinates, directions, and types
- fill missing data

> The only layer that should know the final JSON/YAML/TOON shape
> is **Modeling**.

---

## Common Extension Patterns

---

## Case 1: Add a new extraction target (e.g., comments)

### Steps

1. **Add an extraction method to Backend**

   ```python
   class Backend(Protocol):
       def extract_comments(self, ...): ...
   ```

2. Implement in `OpenpyxlBackend` / `ComBackend`
   - Either side can be optional; use `NotImplementedError` when missing

3. Add calls in `pipeline.py`
   - Make it explicit whether it participates in fallback

4. Integrate into WorkbookData in `modeling.py`

5. Add tests

---

## Case 2: Add a new Backend (e.g., XML or LibreOffice backend)

### Steps

1. Implement the Protocol in `backend.py`

   ```python
   class XmlBackend:
       def extract_cells(...)
       def extract_shapes(...)
   ```

2. Add backend selection in Pipeline
   - Keep changes to existing backends minimal

3. Keep Modeling unchanged if possible

---

## Case 3: Change output structure

- **This is the most fragile type of change**

### Principles

- Limit changes to `modeling.py` and Pydantic models
- Do not change backends
- Do not modify Pipeline

---

## Fallback Rules

- COM or LibreOffice runtime being unavailable is **normal**
- fallback is not an exception
- always provide a `FallbackReason`

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

## Testing Guidelines (Important)

### Expected test granularity

| Layer    | Test focus             |
| -------- | ---------------------- |
| Backend  | extraction correctness |
| Pipeline | fallback / branching   |
| Modeling | integration logic      |

### Anti-patterns

- Fragile tests that depend heavily on a real Excel instance
- Giant tests that combine Backend and Modeling

---

## Pre-PR Checklist

- [ ] No Excel parsing logic in Pipeline
- [ ] No interpretation logic in Backend
- [ ] Modeling is the single source of final structure
- [ ] Fallback reasons are explicit
- [ ] Tests are added
- [ ] Docs updated if public API changes

---

## Common Anti-Patterns

- Creating WorkbookData inside Backend
- Calling openpyxl / xlwings directly in Pipeline
- "Just process here" ad-hoc logic
- Catch-all exceptions without fallback reason

---

## Design Philosophy (Recap)

- Excel is **fragile**
- COM is **powerful but unstable**
- LLM/RAG requires **stable structure first**

Therefore:

> "Separate responsibilities and localize failure points."
