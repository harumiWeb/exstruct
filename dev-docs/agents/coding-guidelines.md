# ExStruct Coding Guidelines

These are ExStruct-specific coding conventions for achieving a codebase that **AI (Codex) and humans can maintain together over the long term**.
The conventions below also apply when having Codex generate code.
They are intended to stabilize the quality of AI-generated code, improve maintainability, and align with Ruff / mypy / Pydantic.

---

# 1. Core Principles (Most Important)

## 1.1 Type hints are required

Add complete type hints to all functions, methods, and classes.

**Good**

```python
def extract_shapes(sheet: xw.Sheet) -> list[Shape]:
    ...
```

**Bad**

```python
def extract_shapes(sheet):
    ...
```

---

## 1.2 Functions should have a single responsibility

Because complex Excel processing tends to bloat functions,
enforce **1 function = 1 logic** strictly.

**Good examples**

- `extract_raw_shapes`
- `normalize_shape`
- `detect_shape_direction`

---

## 1.3 Use BaseModel at boundaries, dataclass internally as primary data structures

Raw*Data (e.g., `SheetRawData`, `WorkbookRawData`) are for internal use only and must not appear in the public API
or be re-exported from the package `__init__`.

Always return a **Pydantic `BaseModel` or dataclass**, never a dictionary or tuple.

Benefits:
- Less ambiguous for AI
- Easier to develop with an IDE
- Better compatibility with JSON serialization

---

# 2. Naming Conventions

## 2.1 snake_case (functions and variables)

Examples: `parse_chart_labels`, `shape_items`

## 2.2 PascalCase (classes)

Examples: `WorkbookParser`, `ChartSeries`

## 2.3 Module names should be short and indicate responsibility

Examples: `shape_parser.py`, `chart_reader.py`

---

# 3. Import Rules

Follow Ruff's isort rules.

Order:

1. Standard library
2. Third-party
3. Internal library (exstruct)

**Example**

```python
import json
from typing import Any

import xlwings as xw
from pydantic import BaseModel

from exstruct.models import Shape
```

---

# 4. Docstring Conventions

Use Google style and always include:

- Args
- Returns
- Raises (if the function raises exceptions)

**Example**

```python
def detect_shape_direction(shape: Shape) -> str | None:
    """Detect arrow direction from shape coordinates and rotation.

    Args:
        shape: Parsed shape model.

    Returns:
        Direction code ("E", "NE", etc.) or None if no arrow is detected.
    """
```

---

# 5. Exception Handling Rules

## 5.1 Use ValueError / RuntimeError as the primary exception types

```python
if not cell:
    raise ValueError("cell must not be empty.")
```

## 5.2 Wrap COM exceptions

```python
try:
    text = shape.text_frame.characters.text
except Exception as e:
    raise RuntimeError(f"Failed to read shape text: {e}") from e
```

---

# 6. Complexity Control

Following Ruff's `C90` (mccabe), split functions so that
**max-complexity = 12** is never exceeded.

---

# 7. Guidelines for Using Codex (AI)

## 7.1 Rules to enforce with Codex

When using Codex, include the following in the prompt in advance:

- Type hints are required
- 1 function = 1 responsibility
- Return BaseModel at boundaries, dataclass internally
- Write docstrings in Google style
- Write imports in the correct order
- Keep error handling concise
- Split functions when complexity gets too high

---

## 7.2 Recommended prompt for Codex output

Paste the following for stable code generation:

```text
You are an experienced Python engineer who writes code following the ExStruct library standards.

Always follow these rules:
- Add type hints to all arguments and return values
- 1 function = 1 responsibility
- Return BaseModel at boundaries, dataclass internally
- Order imports correctly
- Write docstrings (Google style)
- Split functions to avoid getting too complex
- Return Pydantic models, not JSON or dictionaries

Output Python code only.
```

---

# 8. Review Checklist for AI-Generated Code

When reviewing AI-generated code, check the following in order:

1. Are type hints complete?
2. Are docstrings present?
3. Is the import order correct?
4. Does it return BaseModel at boundaries and dataclass internally?
5. Does each function have a single responsibility?
6. Is exception handling appropriate?
7. Does complexity not exceed max 12?
8. Does Ruff report no errors?

---

# 9. Prohibited Patterns (Especially for AI-Generated Code)

Never let Codex do the following:

- Overly complex if/else nesting
- God classes with many responsibilities
- Returning large dictionaries or tuples
- Code with no comments at all
- Anonymous magic numbers

---

# 10. Summary of ExStruct Design Principles

- Type safety
- Clear data structures (BaseModel at boundaries, dataclass internally)
- Loose coupling between modules
- Minimal function responsibility
- Code discipline that assumes AI and humans work side by side

These are designed to maintain **high compatibility with Ruff / mypy / CI**.

---

# Appendix: Rules Recommended for Future Addition

- Stricter mypy enforcement
- Unified Pydantic field constraints
- Mandatory docstrings for public API
- `_prefix` naming for internal APIs
