````md
**Actionable comments posted: 0**

> [!CAUTION]
> Some comments are outside the diff and can’t be posted inline due to platform limitations.
>
> <details>
> <summary>⚠️ Outside diff range comments (1)</summary><blockquote>
>
> <details>
> <summary>src/exstruct/io/__init__.py (1)</summary><blockquote>
>
> `74-82`: **Return `RangeBounds` directly instead of converting to tuple.**
>
> This wrapper function unpacks the Pydantic `RangeBounds` model into a tuple, which violates the coding guideline: "Do not return dictionaries or tuples; always use Pydantic BaseModel for structured data." Callers should access bounds via the model's fields (`bounds.r1`, `bounds.c1`, etc.) to preserve type safety and semantic clarity.
>
> As per coding guidelines, structured data should be returned as Pydantic models, not tuples.
>
> <details>
> <summary>🔎 Proposed refactor to eliminate tuple conversion</summary>
>
> **Option 1: Remove the wrapper entirely and use `parse_range_zero_based` directly**
>
> Update callers (e.g., line 129) to use the model fields:
>
> ```diff
>  def _filter_table_candidates_to_area(
>      table_candidates: list[str], area: PrintArea
>  ) -> list[str]:
>      filtered: list[str] = []
>      for candidate in table_candidates:
> -        bounds = _parse_range_zero_based(candidate)
> -        if not bounds:
> +        bounds = parse_range_zero_based(candidate)
> +        if bounds is None:
>              continue
> -        r1, c1, r2, c2 = bounds
> +        r1, c1, r2, c2 = bounds.r1, bounds.c1, bounds.r2, bounds.c2
>          r1 += 1
>          r2 += 1
>          if r1 >= area.r1 and r2 <= area.r2 and c1 >= area.c1 and c2 <= area.c2:
>              filtered.append(candidate)
>      return filtered
> ```
>
> **Option 2: If the wrapper is needed, change return type to `RangeBounds | None`**
>
> ```diff
> -def _parse_range_zero_based(range_str: str) -> tuple[int, int, int, int] | None:
> -    """
> -    Parse an Excel range string into zero-based (r1, c1, r2, c2) bounds.
> -    Returns None on failure.
> -    """
> +def _parse_range_zero_based(range_str: str) -> RangeBounds | None:
> +    """Parse an Excel range string into zero-based bounds.
> +
> +    Args:
> +        range_str: Excel range string (e.g., "Sheet1!A1:B2").
> +
> +    Returns:
> +        RangeBounds in zero-based coordinates, or None on failure.
> +    """
>      bounds = parse_range_zero_based(range_str)
> -    if bounds is None:
> -        return None
> -    return (bounds.r1, bounds.c1, bounds.r2, bounds.c2)
> +    return bounds
> ```
>
> </details>
>
> </blockquote></details>
>
> </blockquote></details>

<details>
<summary>♻️ Duplicate comments (3)</summary><blockquote>

<details>
<summary>src/exstruct/core/pipeline.py (2)</summary><blockquote>

`644-645`: **Previous review concern addressed.**

The error message has been updated to use a generic "COM pipeline failed" message with `FallbackReason.COM_PIPELINE_FAILED`, addressing the earlier review comment about misleading error messages.

---

`670-670`: **Previous review concern addressed.**

The `reason` parameter is now properly used for logging, addressing the earlier review comment about the unused parameter.

</blockquote></details>
<details>
<summary>docs/agents/TEST_REQUIREMENTS.md (1)</summary><blockquote>

`203-204`: **Previous review concern addressed.**

The duplicate requirement IDs have been resolved by renaming them to `[PIPE-MOD-01]` and `[PIPE-MOD-02]`, making each requirement ID unique throughout the document.

</blockquote></details>

</blockquote></details>

<details>
<summary>🧹 Nitpick comments (7)</summary><blockquote>

<details>
<summary>src/exstruct/io/__init__.py (1)</summary><blockquote>

`488-489`: **Consider renaming publicly exported functions to remove underscore prefix.**

The functions `_require_yaml` and `_require_toon` are now part of the public API (exported in `__all__`), but their names start with an underscore, which by Python convention indicates private/internal functions. For a cleaner public API, consider renaming them to `require_yaml` and `require_toon` (without the underscore).

<details>
<summary>🔎 Suggested changes for consistent public API naming</summary>

In `src/exstruct/io/serialize.py`, rename the functions:

```diff
-def _require_yaml() -> ModuleType:
+def require_yaml() -> ModuleType:
     """Ensure pyyaml is installed; otherwise raise with guidance."""
     ...

-def _require_toon() -> ModuleType:
+def require_toon() -> ModuleType:
     """Ensure python-toon is installed; otherwise raise with guidance."""
     ...
```

In this file, update the import and export:

```diff
 from .serialize import (
     _FORMAT_HINTS,
     _ensure_format_hint,
-    _require_toon,
-    _require_yaml,
+    require_toon,
+    require_yaml,
     _serialize_payload_from_hint,
 )

 __all__ = [
     "dict_without_empty_values",
     "save_as_json",
     "save_as_yaml",
     "save_as_toon",
     "save_sheets",
     "save_sheets_as_json",
     "build_print_area_views",
     "save_print_area_views",
     "save_auto_page_break_views",
     "serialize_workbook",
-    "_require_yaml",
-    "_require_toon",
+    "require_yaml",
+    "require_toon",
 ]
```

</details>

</blockquote></details>
<details>
<summary>tests/test_backends.py (3)</summary><blockquote>

`16-22`: **Add type hints to mock functions.**

The fake functions should have explicit type hints for maintainability and mypy compliance.

<details>
<summary>🔎 Proposed refactor</summary>

```diff
-    def fake_cells(_: Path) -> dict[str, list[object]]:
+    def fake_cells(file_path: Path) -> dict[str, list[object]]:
         calls.append("cells")
         return {}

-    def fake_cells_links(_: Path) -> dict[str, list[object]]:
+    def fake_cells_links(file_path: Path) -> dict[str, list[object]]:
         calls.append("links")
         return {}
```

</details>

As per coding guidelines, avoid using `_` for actual parameters; use descriptive names with proper type hints.

---

`43-44`: **Use explicit parameter names with type hints.**

Replace generic `_` and `__` with descriptive parameter names for better readability.

<details>
<summary>🔎 Proposed refactor</summary>

```diff
-    def fake_detect(_: Path, __: str) -> list[str]:
+    def fake_detect(file_path: Path, sheet_name: str) -> list[str]:
         raise RuntimeError("boom")
```

</details>

As per coding guidelines, use descriptive parameter names.

---

`58-59`: **Use explicit type hints instead of generic object.**

The mock function should use proper type signatures for clarity.

<details>
<summary>🔎 Proposed refactor</summary>

```diff
-    def fake_colors_map(*_: object, **__: object) -> object:
+    def fake_colors_map(
+        workbook: object,
+        *,
+        include_default_background: bool,
+        ignore_colors: set[str] | None
+    ) -> object:
         raise RuntimeError("boom")
```

</details>

As per coding guidelines, provide explicit type hints for all parameters.

</blockquote></details>
<details>
<summary>src/exstruct/core/backends/openpyxl_backend.py (1)</summary><blockquote>

`103-125`: **Consider adding a module-level docstring.**

The file contains well-structured code with proper docstrings for classes and methods, but lacks a module-level docstring explaining the openpyxl backend's role in the extraction pipeline.

<details>
<summary>🔎 Suggested addition</summary>

Add at the top of the file after imports:

```python
"""Openpyxl-based backend for Excel workbook extraction.

This module provides the OpenpyxlBackend class which uses the openpyxl library
to extract cells, print areas, color maps, and table candidates from Excel files.
It serves as the primary extraction backend with fallback support when COM is unavailable.
"""
```

</details>

As per coding guidelines, adding comprehensive module documentation improves maintainability.

</blockquote></details>
<details>
<summary>src/exstruct/core/backends/com_backend.py (2)</summary><blockquote>

`173-206`: **Consider extracting CSV parsing to a utility function for reusability.**

The `_split_csv_respecting_quotes` function implements custom CSV parsing with quote handling. While the implementation is correct, this type of utility might be useful elsewhere in the codebase.

Consider moving this to a shared utilities module if similar parsing is needed elsewhere, or documenting that Python's `csv` module with appropriate dialect settings could be an alternative.

---

`1-206`: **Add module-level docstring for better documentation.**

The file contains well-structured code with proper class and method docstrings, but lacks a module-level docstring explaining the COM backend's role.

<details>
<summary>🔎 Suggested addition</summary>

Add at the top after imports:

```python
"""COM-based backend for Excel workbook extraction via xlwings.

This module provides the ComBackend class which uses xlwings/COM to extract
print areas, color maps, and auto page breaks from Excel files. It complements
the OpenpyxlBackend and is used when COM automation is available and required
features (like auto page breaks) need native Excel API access.
"""
```

</details>

As per coding guidelines, comprehensive module documentation improves maintainability.

</blockquote></details>

</blockquote></details>

<details>
<summary>📜 Review details</summary>

**Configuration used**: defaults

**Review profile**: CHILL

**Plan**: Pro

<details>
<summary>📥 Commits</summary>

Reviewing files that changed from the base of the PR and between 79d4bbb797caaa2641a3581e2d56059fc5a8cbf2 and 93b10e2d8b0e9e4e339f038cb215951dd4fdd7c2.

</details>

<details>
<summary>📒 Files selected for processing (20)</summary>

- `.github/workflows/pytest.yml`
- `docs/agents/CODE_REVIEW.md`
- `docs/agents/DATA_MODEL.md`
- `docs/agents/EXCEL_EXTRACTION.md`
- `docs/agents/TEST_REQUIREMENTS.md`
- `src/exstruct/core/backends/com_backend.py`
- `src/exstruct/core/backends/openpyxl_backend.py`
- `src/exstruct/core/pipeline.py`
- `src/exstruct/errors.py`
- `src/exstruct/io/__init__.py`
- `src/exstruct/models/__init__.py`
- `tests/test_auto_page_breaks.py`
- `tests/test_backends.py`
- `tests/test_engine.py`
- `tests/test_integrate_raw_data.py`
- `tests/test_modeling.py`
- `tests/test_models_validation.py`
- `tests/test_pipeline.py`
- `tests/test_print_area_views.py`
- `tests/test_print_areas_openpyxl.py`

</details>

<details>
<summary>🚧 Files skipped from review as they are similar to previous changes (4)</summary>

- tests/test_pipeline.py
- tests/test_modeling.py
- tests/test_integrate_raw_data.py
- src/exstruct/errors.py

</details>

<details>
<summary>🧰 Additional context used</summary>

<details>
<summary>📓 Path-based instructions (1)</summary>

<details>
<summary>**/*.py</summary>

**📄 CodeRabbit inference engine (AGENTS.md)**

> `**/*.py`: Add type hints to all function and method arguments and return values (mypy strict compliance)
> Use `Any` type only at external library boundaries (xlwings, pandas, numpy, etc.)
> Return Pydantic BaseModel instances instead of dictionaries or tuples for structured data
> Ensure each function has a single responsibility (cyclomatic complexity should not exceed 12)
> Organize imports in this order: (1) standard library, (2) third-party packages, (3) exstruct internal modules
> Use Google-style docstrings for all functions and classes
> Do not depend on internal structures of external libraries (xlwings, pandas, numpy); normalize external data to Pydantic models at boundaries
> Avoid writing God Functions (large, single-responsibility-violating functions)
> Avoid writing God Objects (classes with too many responsibilities)
> Avoid excessive nesting and deep conditional branching
> Do not return dictionaries or tuples; always use Pydantic BaseModel for structured data
> Ensure code passes mypy strict mode with zero errors
> Ensure code passes Ruff checks (E, W, F, I, B, UP, N, C90) with zero errors
> Avoid circular dependencies between modules
> AI should propose automatic function splitting when functions are too complex
> AI should propose additional Pydantic models when data structures are insufficient
> AI should automatically fix and organize imports when they are inconsistent
> AI should improve docstrings when they are incomplete or insufficient

Files:

- `tests/test_print_areas_openpyxl.py`
- `tests/test_backends.py`
- `src/exstruct/core/backends/openpyxl_backend.py`
- `src/exstruct/io/__init__.py`
- `src/exstruct/core/backends/com_backend.py`
- `tests/test_print_area_views.py`
- `tests/test_engine.py`
- `tests/test_auto_page_breaks.py`
- `src/exstruct/core/pipeline.py`
- `tests/test_models_validation.py`
- `src/exstruct/models/__init__.py`

</details>

</details><details>
<summary>🧠 Learnings (1)</summary>

<details>
<summary>📚 Learning: 2025-12-21T10:55:22.238Z</summary>

```
Learnt from: CR
Repo: harumiWeb/exstruct PR: 0
File: AGENTS.md:0-0
Timestamp: 2025-12-21T10:55:22.238Z
Learning: Applies to **/*.py : Do not depend on internal structures of external libraries (xlwings, pandas, numpy); normalize external data to Pydantic models at boundaries
```

**Applied to files:**

- `src/exstruct/io/__init__.py`

</details>

</details><details>
<summary>🧬 Code graph analysis (6)</summary>

<details>
<summary>tests/test_backends.py (3)</summary><blockquote>

<details>
<summary>src/exstruct/core/backends/com_backend.py (3)</summary>

- `ComBackend` (18-135)
- `extract_colors_map` (51-74)
- `extract_print_areas` (27-49)

</details>
<details>
<summary>src/exstruct/core/backends/openpyxl_backend.py (5)</summary>

- `OpenpyxlBackend` (23-100)
- `extract_cells` (32-45)
- `detect_tables` (88-100)
- `extract_colors_map` (64-86)
- `extract_print_areas` (47-62)

</details>
<details>
<summary>src/exstruct/core/ranges.py (1)</summary>

- `parse_range_zero_based` (24-47)

</details>

</blockquote></details>
<details>
<summary>src/exstruct/core/backends/openpyxl_backend.py (3)</summary><blockquote>

<details>
<summary>src/exstruct/models/__init__.py (1)</summary>

- `PrintArea` (97-103)

</details>
<details>
<summary>src/exstruct/core/ranges.py (1)</summary>

- `parse_range_zero_based` (24-47)

</details>
<details>
<summary>src/exstruct/core/workbook.py (1)</summary>

- `openpyxl_workbook` (14-53)

</details>

</blockquote></details>
<details>
<summary>src/exstruct/io/__init__.py (3)</summary><blockquote>

<details>
<summary>src/exstruct/core/ranges.py (1)</summary>

- `parse_range_zero_based` (24-47)

</details>
<details>
<summary>src/exstruct/errors.py (2)</summary>

- `OutputError` (32-33)
- `SerializationError` (20-21)

</details>
<details>
<summary>src/exstruct/io/serialize.py (4)</summary>

- `_ensure_format_hint` (28-49)
- `_require_toon` (104-112)
- `_require_yaml` (93-101)
- `_serialize_payload_from_hint` (52-90)

</details>

</blockquote></details>
<details>
<summary>tests/test_print_area_views.py (1)</summary><blockquote>

<details>
<summary>src/exstruct/models/__init__.py (2)</summary>

- `CellRow` (50-59)
- `PrintArea` (97-103)

</details>

</blockquote></details>
<details>
<summary>tests/test_auto_page_breaks.py (1)</summary><blockquote>

<details>
<summary>src/exstruct/models/__init__.py (1)</summary>

- `PrintArea` (97-103)

</details>

</blockquote></details>
<details>
<summary>tests/test_models_validation.py (1)</summary><blockquote>

<details>
<summary>src/exstruct/models/__init__.py (1)</summary>

- `CellRow` (50-59)

</details>

</blockquote></details>

</details><details>
<summary>🪛 GitHub Check: Codacy Static Code Analysis</summary>

<details>
<summary>src/exstruct/core/backends/com_backend.py</summary>

[warning] 37-37: src/exstruct/core/backends/com_backend.py#L37
Try, Except, Continue detected.

---

[warning] 132-132: src/exstruct/core/backends/com_backend.py#L132
Try, Except, Pass detected.

</details>

</details>

</details>

<details>
<summary>⏰ Context from checks skipped due to timeout of 90000ms. You can increase the timeout in your CodeRabbit configuration to a maximum of 15 minutes (900000ms). (2)</summary>

- GitHub Check: Codacy Static Code Analysis
- GitHub Check: test (windows-latest, 3.11)

</details>

<details>
<summary>🔇 Additional comments (24)</summary><blockquote>

<details>
<summary>docs/agents/DATA_MODEL.md (2)</summary><blockquote>

`96-96`: **Clear coordinate semantics for PrintArea rows—good clarification.**

Adding "(1-based, inclusive)" to `r1` and `r2` aligns with the existing `c1` and `c2` documentation and makes the row coordinate interpretation explicit. This improves clarity for API consumers.

Also applies to: 98-98

---

`140-140`: **Documentation for `colors_map` coordinate system is correct and matches implementation.**

The codebase confirms that `colors_map` uses (row=1-based, col=0-based) coordinates as documented. Both extraction paths—openpyxl (`cell.row, cell.col_idx - 1`) and COM (`row, col - 1`)—apply the same logic: rows are preserved as 1-based, columns are converted from 1-based to 0-based. Test coverage validates this with `{"#FFFFFF": [(1, 0)]}`, confirming row 1 with column 0.

</blockquote></details>
<details>
<summary>src/exstruct/io/__init__.py (3)</summary><blockquote>

`8-18`: **Well done removing the OpenPyXL dependency!**

The refactor successfully eliminates direct dependency on `openpyxl.utils.cell.range_boundaries` by using the internal `parse_range_zero_based` function, and centralizes serialization logic through the new helpers from `.serialize`. This aligns with the coding guideline to avoid depending on internal structures of external libraries.

Based on learnings, this change properly normalizes external data to Pydantic models at boundaries.

---

`283-288`: **Excellent serialization refactoring!**

The centralization of format validation and serialization through `_ensure_format_hint` and `_serialize_payload_from_hint` eliminates code duplication and provides a single source of truth for format handling. This improves maintainability and consistency across all export functions.

Also applies to: 315-318, 341-346, 374-377, 393-402, 429-432, 449-454, 469-471

---

`133-134`: **Row index adjustments correctly implement 1-based row coordinates.**

The adjustments on lines 133-134 (adding 1 to `r1` and `r2`) and lines 148, 150 (subtracting 1 from `area.r1` in pixel calculations) correctly convert between the 0-based coordinates returned by `parse_range_zero_based` and the 1-based row coordinates used in `PrintArea`, as documented in the PR objectives.

Also applies to: 148-148, 150-150

</blockquote></details>
<details>
<summary>tests/test_models_validation.py (1)</summary><blockquote>

`19-19`: **LGTM: Coordinate system alignment.**

The update from `r=0` to `r=1` correctly reflects the 1-based row indexing documented in the `CellRow` model definition.

</blockquote></details>
<details>
<summary>tests/test_auto_page_breaks.py (1)</summary><blockquote>

`47-47`: **LGTM: Coordinate system alignment.**

The updated PrintArea coordinates correctly reflect the 1-based row indexing (r1=1, r2=2) and 0-based column indexing (c1=0, c2=1) documented in the PrintArea model.

</blockquote></details>
<details>
<summary>.github/workflows/pytest.yml (1)</summary><blockquote>

`34-46`: **LGTM: Coverage collection and Codecov integration properly configured.**

The pytest coverage flags and Codecov upload step are correctly implemented. The Codecov action is now properly pinned to a full commit SHA, addressing the previous security concern.

</blockquote></details>
<details>
<summary>src/exstruct/models/__init__.py (3)</summary><blockquote>

`14-16`: **LGTM: Field definition reflow.**

The Shape.id field definition has been reflowed for readability while maintaining the same semantic meaning.

---

`101-103`: **LGTM: Coordinate system documentation clarified.**

The field descriptions now explicitly document that columns use 0-based indexing, improving clarity for API consumers.

---

`129-132`: **LGTM: colors_map coordinate semantics documented.**

The description now explicitly clarifies that rows are 1-based and columns are 0-based, which is essential for correct usage of the colors_map field.

</blockquote></details>
<details>
<summary>docs/agents/EXCEL_EXTRACTION.md (1)</summary><blockquote>

`1-66`: **LGTM: Comprehensive pipeline documentation.**

This documentation provides clear guidance on the extraction pipeline architecture, coordinate system semantics, modes, and fallback behavior. The coordinate system clarification (rows 1-based, columns 0-based) is consistent with the code changes throughout the PR.

</blockquote></details>
<details>
<summary>src/exstruct/core/pipeline.py (1)</summary><blockquote>

`1-696`: **LGTM: Pipeline architecture with proper type safety.**

The pipeline module demonstrates excellent adherence to coding guidelines:

- Complete type hints on all functions and parameters
- Google-style docstrings throughout
- Immutable dataclasses for pipeline configuration and state
- Well-organized imports (stdlib → third-party → internal)
- Clear separation of pre-COM and COM extraction steps

The architecture provides a solid foundation for the extraction workflow with explicit fallback handling and state tracking.

</blockquote></details>
<details>
<summary>docs/agents/TEST_REQUIREMENTS.md (1)</summary><blockquote>

`1-228`: **LGTM: Comprehensive test requirements specification.**

The test requirements document provides thorough coverage of functional, non-functional, and integration requirements. The organization by category (pipeline, backend, ranges, etc.) aligns well with the modular architecture introduced in this PR.

</blockquote></details>
<details>
<summary>tests/test_print_areas_openpyxl.py (1)</summary><blockquote>

`26-26`: **LGTM: Coordinate system alignment.**

The assertion now correctly expects 1-based row coordinates (r1=1, r2=2) and 0-based column coordinates (c1=0, c2=1), consistent with the PrintArea model definition and the broader coordinate system updates in this PR.

</blockquote></details>
<details>
<summary>tests/test_print_area_views.py (1)</summary><blockquote>

`39-46`: **LGTM! Coordinate system update correctly applied.**

The test data has been properly updated to reflect the documented coordinate convention (1-based rows, 0-based columns). The CellRow indices and PrintArea bounds are now consistent with the model definitions.

</blockquote></details>
<details>
<summary>tests/test_backends.py (2)</summary><blockquote>

`76-93`: **LGTM! Print area extraction test validates coordinate system.**

The test correctly verifies that print areas extracted via openpyxl use 1-based row indexing (r1=1) and 0-based column indexing (c1=0), consistent with the PrintArea model definition.

---

`95-101`: **LGTM! Range parsing test validates zero-based intermediate representation.**

The test correctly verifies that `parse_range_zero_based` returns zero-based coordinates (r1=0, c1=0, r2=1, c2=1) for the range "Sheet1!A1:B2", which are then converted to the appropriate coordinate system by backend methods.

</blockquote></details>
<details>
<summary>src/exstruct/core/backends/openpyxl_backend.py (1)</summary><blockquote>

`159-162`: **LGTM! Coordinate conversion correctly implemented.**

The conversion from zero-based coordinates (from `_parse_print_area_range`) to the PrintArea model is now correct:

- Rows (r1, r2): converted from 0-based to 1-based with `+ 1`
- Columns (c1, c2): kept as 0-based (no adjustment)

This matches the PrintArea model definition where rows are 1-based and columns are 0-based.

</blockquote></details>
<details>
<summary>tests/test_engine.py (2)</summary><blockquote>

`6-12`: **LGTM! Improved API design with nested option groups.**

The introduction of `FilterOptions` and `DestinationOptions` improves the API's organization and clarity, grouping related settings together under `OutputOptions`.

---

`65-65`: **LGTM! PrintArea coordinate updated correctly.**

The PrintArea coordinates have been updated to use 1-based row indexing (r1=1) consistent with the model definition and broader coordinate system clarification in this PR.

</blockquote></details>
<details>
<summary>docs/agents/CODE_REVIEW.md (1)</summary><blockquote>

`1-609`: **Documentation file - no code review needed.**

This file contains example code review comments and AI agent prompts for documentation purposes. It does not contain production code requiring review.

</blockquote></details>
<details>
<summary>src/exstruct/core/backends/com_backend.py (2)</summary><blockquote>

`45-48`: **LGTM! Coordinate conversion correctly implemented.**

The conversion from zero-based coordinates (from `_parse_print_area_range`) to PrintArea is correct:

- Rows (r1, r2): `+ 1` converts 0-based to 1-based
- Columns (c1, c2): no adjustment keeps 0-based

This matches the PrintArea model definition.

---

`119-125`: **LGTM! Coordinate conversion for COM values is correct.**

The coordinate handling here is actually correct, contrary to the past review comment:

- `rows[i]` and `cols[j]` are 1-based values from Excel COM
- Rows (r1, r2): used as-is since PrintArea expects 1-based rows ✓
- Columns (c1, c2): `- 1` converts from 1-based COM to 0-based for PrintArea ✓

This is consistent with the PrintArea model where rows are 1-based and columns are 0-based.

</blockquote></details>

</blockquote></details>

</details>

<!-- This is an auto-generated comment by CodeRabbit for review status -->
````
