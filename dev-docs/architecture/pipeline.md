# Pipeline Architecture Overview

ExStruct uses a three-layer **Pipeline + Backend + Modeling** architecture
to convert Excel workbooks into **semantically structured JSON**.

This design achieves the following.

- Separation of Excel COM-dependent logic from non-dependent logic
- Direct OOXML reading for the normal light path on .xlsx / .xlsm
- Stable output for RAG/LLM use cases

---

## End-to-End Flow

```mermaid
sequenceDiagram
    participant Client
    participant Pipeline
    participant OoxmlSession
    participant OpenpyxlBackend
    participant RichBackend
    participant Modeling

    Client->>Pipeline: extract()
    alt supported light .xlsx/.xlsm
        Pipeline->>OoxmlSession: core + rich extraction (one ZIP)
        OoxmlSession-->>Pipeline: cells / tables / print_areas / shapes / charts
        opt unsupported construct or direct-stage failure
            Pipeline->>Pipeline: discard partial result; close session
            Pipeline->>OpenpyxlBackend: restart complete compatibility pipeline
            OpenpyxlBackend-->>Pipeline: complete compatibility result
        end
    else other modes, .xls, or compatibility selection
        Pipeline->>OpenpyxlBackend: selected extraction path
        OpenpyxlBackend-->>Pipeline: core extraction
        Pipeline->>RichBackend: selected rich extraction
        RichBackend-->>Pipeline: rich artifacts
    end

    Pipeline->>Modeling: integrate()
    Modeling-->>Pipeline: WorkbookData
    Pipeline-->>Client: structured output
```

The processing order is as follows.

`RichBackend` in this diagram refers to the conceptual rich-extraction layer; the concrete implementations are `OoxmlRichBackend`, `ComRichBackend`, and `LibreOfficeRichBackend`.

1. **Pipeline** selects the mode and backend.
2. Supported .xlsx / .xlsm light extraction reads core and rich OOXML data
   through one OoxmlExtractionSession ZIP.
3. An unsupported construct or uncaught failure at any direct stage discards
   the partial result, closes the session, and restarts the full openpyxl
   compatibility pipeline. colors_map opt-in and active legacy overrides
   select compatibility directly.
4. Other modes and .xls retain their existing backend selection. The
   conceptual rich layer includes OoxmlRichBackend, ComRichBackend, and
   LibreOfficeRichBackend.
5. **Modeling** integrates the result into WorkbookData / SheetData; output is
   serialized in the requested format (JSON / YAML / TOON).

---

## Pipeline Responsibilities

Pipeline is the **orchestrator**.

- Determines the extraction order
- Selects backends
- Controls fallback paths
- Manages intermediate artifacts
- Owns one extraction-scoped OOXML session for normal light .xlsx / .xlsm
  processing, or one openpyxl session when the compatibility/other path is
  selected. Resources remain scoped through final model construction.

Pipeline is designed to **never read Excel content directly**.

OoxmlExtractionSession shares one ZIP between core and rich light extraction.
When openpyxl compatibility is selected, OpenpyxlExtractionSession owns the
regular workbook variants: cached values (data_only=True) are shared by
compatible stages, and formula text (data_only=False) is opened only when
requested. Neither session is a global cache or outlives one extraction.

---

## Backend Responsibilities

Backend defines **how Excel is read**.

| Backend                | Responsibilities                                  |
| ---------------------- | ------------------------------------------------- |
| OpenpyxlBackend        | Cells / tables / print areas / colors map         |
| OoxmlExtractionSession | Light-mode OOXML core data and shared ZIP lifetime |
| ComBackend             | COM-only print areas / auto page breaks / maps    |
| OoxmlRichBackend       | Pure-Python OOXML shapes / connectors / charts    |
| ComRichBackend         | Shapes / arrows / charts / SmartArt via Excel COM |
| LibreOfficeRichBackend | LibreOffice-enriched shapes / connectors / charts |

In this document, `RichBackend` refers to the protocol-level concept, while `OoxmlRichBackend`, `ComRichBackend`, and `LibreOfficeRichBackend` are the concrete backend classes.

This abstraction enables the following extensions.

- Direct XML parsing backend
- LibreOffice backend
- Remote Excel service backend

All of these can be added **without major changes to the Pipeline**.

---

## Fallback Design

Fallback behavior depends on the selected path.

- A direct light OOXML stage failure or unsupported construct discards every
  partial result, closes the session in finally, and restarts the complete
  openpyxl compatibility pipeline with the ooxml_compatibility warning.
- A drawing failure isolated to one sheet retains the existing best-effort
  behavior and does not discard healthy sheet output.
- colors_map opt-in and active legacy overrides select compatibility; other
  modes and .xls retain their existing routing.
- COM/LibreOffice fallback behavior remains as defined by the mode contract.
- Record fallback reasons explicitly through FallbackReason.

This is an intentional design that assumes **batch processing, CI, and automation**.

---

## Modeling Layer Responsibilities

Modeling is responsible for:

- Integrating results from multiple backends
- Producing normalized WorkbookData / SheetData
- Not depending on the output format itself

**Semantic structure models** for RAG/LLM use are centralized here.

---

## Why This Design

- Excel has separate worlds of cells, shapes, and charts
- COM is powerful but fragile
- LLMs require stable structured data

Therefore, **pipeline separation** is the most practical approach.
